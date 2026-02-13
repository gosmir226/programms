"""
Унифицированный модуль управления кэшем для всех программ.

Кэш хранит информацию о обработанных файлах с их хэшами и датами модификации
для быстрой проверки необходимости повторной обработки.
"""

import os
import json
import hashlib
from datetime import datetime
from typing import Dict, List, Optional, Tuple


class CacheManager:
    """Менеджер кэша для отслеживания обработанных файлов."""
    
    def __init__(self, program_name: str, output_file_path: str, cache_dir: Optional[str] = None):
        """
        Инициализация менеджера кэша.
        
        Args:
            program_name: Название программы (для идентификации)
            output_file_path: Путь к результирующему файлу
            cache_dir: Директория для хранения кэша (по умолчанию - рядом с output файлом)
        """
        self.program_name = program_name
        self.output_file_path = os.path.abspath(output_file_path)
        
        # Определяем путь к кэш-файлу
        if cache_dir:
            self.cache_dir = cache_dir
        else:
            self.cache_dir = os.path.dirname(self.output_file_path)
        
        # Создаем уникальное имя кэш-файла на основе имени output файла
        output_basename = os.path.splitext(os.path.basename(self.output_file_path))[0]
        self.cache_file = os.path.join(
            self.cache_dir, 
            f'.cache_{program_name}_{output_basename}.json'
        )
        
        self.cache_data = self._load_cache()
    
    def _load_cache(self) -> dict:
        """Загружает кэш из файла."""
        if not os.path.exists(self.cache_file):
            return self._create_empty_cache()
        
        try:
            with open(self.cache_file, 'r', encoding='utf-8') as f:
                cache = json.load(f)
                
            # Проверяем валидность кэша
            if cache.get('program_name') != self.program_name:
                print(f"[КЭШ] Кэш создан другой программой, создается новый")
                return self._create_empty_cache()
            
            if cache.get('output_file') != self.output_file_path:
                print(f"[КЭШ] Кэш для другого output файла, создается новый")
                return self._create_empty_cache()
            
            return cache
            
        except (json.JSONDecodeError, KeyError) as e:
            print(f"[КЭШ] Ошибка чтения кэша: {e}. Создается новый.")
            return self._create_empty_cache()
    
    def _create_empty_cache(self) -> dict:
        """Создает пустую структуру кэша."""
        return {
            'program_name': self.program_name,
            'output_file': self.output_file_path,
            'input_files': {},
            'last_update': None
        }
    
    def _save_cache(self) -> bool:
        """Сохраняет кэш в файл."""
        try:
            self.cache_data['last_update'] = datetime.now().isoformat()
            
            # Создаем директорию если не существует
            os.makedirs(self.cache_dir, exist_ok=True)
            
            with open(self.cache_file, 'w', encoding='utf-8') as f:
                json.dump(self.cache_data, f, ensure_ascii=False, indent=2)
            
            return True
            
        except Exception as e:
            print(f"[КЭШ] Ошибка сохранения: {e}")
            return False
    
    def _calculate_file_hash(self, file_path: str) -> Optional[str]:
        """
        Вычисляет SHA-256 хэш файла.
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            str: Хэш файла или None при ошибке
        """
        try:
            hash_sha256 = hashlib.sha256()
            with open(file_path, 'rb') as f:
                for chunk in iter(lambda: f.read(8192), b''):
                    hash_sha256.update(chunk)
            return hash_sha256.hexdigest()
        except Exception as e:
            print(f"[КЭШ] Ошибка вычисления хэша для {file_path}: {e}")
            return None
    
    def is_file_changed(self, file_path: str) -> bool:
        """
        Проверяет, изменился ли файл с момента последней обработки.
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            bool: True если файл изменился или не был обработан
        """
        abs_path = os.path.abspath(file_path)
        
        if not os.path.exists(abs_path):
            return False  # Файл не существует
        
        # Файл не в кэше - считаем измененным
        if abs_path not in self.cache_data['input_files']:
            return True
        
        cached_info = self.cache_data['input_files'][abs_path]
        
        # Проверяем дату модификации
        try:
            current_mtime = os.path.getmtime(abs_path)
            if current_mtime != cached_info.get('modified_date'):
                return True
        except OSError:
            return True
        
        # Проверяем размер файла (быстрая проверка)
        try:
            current_size = os.path.getsize(abs_path)
            if current_size != cached_info.get('size'):
                return True
        except OSError:
            return True
        
        # Проверяем хэш (более надежная проверка)
        current_hash = self._calculate_file_hash(abs_path)
        if current_hash != cached_info.get('hash'):
            return True
        
        return False  # Файл не изменился
    
    def update_file(self, file_path: str) -> bool:
        """
        Обновляет информацию о файле в кэше.
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            bool: Успешность операции
        """
        abs_path = os.path.abspath(file_path)
        
        if not os.path.exists(abs_path):
            print(f"[КЭШ] Файл не существует: {abs_path}")
            return False
        
        try:
            file_hash = self._calculate_file_hash(abs_path)
            if file_hash is None:
                return False
            
            self.cache_data['input_files'][abs_path] = {
                'hash': file_hash,
                'modified_date': os.path.getmtime(abs_path),
                'size': os.path.getsize(abs_path),
                'last_processed': datetime.now().isoformat()
            }
            
            return self._save_cache()
            
        except Exception as e:
            print(f"[КЭШ] Ошибка обновления файла в кэше: {e}")
            return False
    
    def remove_file(self, file_path: str) -> bool:
        """
        Удаляет файл из кэша.
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            bool: Успешность операции
        """
        abs_path = os.path.abspath(file_path)
        
        if abs_path in self.cache_data['input_files']:
            del self.cache_data['input_files'][abs_path]
            return self._save_cache()
        
        return True
    
    def get_changed_files(self, file_list: List[str]) -> Tuple[List[str], List[str]]:
        """
        Получает списки измененных и неизмененных файлов.
        
        Args:
            file_list: Список путей к файлам для проверки
            
        Returns:
            Tuple[List[str], List[str]]: (измененные файлы, неизмененные файлы)
        """
        changed = []
        unchanged = []
        
        for file_path in file_list:
            if self.is_file_changed(file_path):
                changed.append(file_path)
            else:
                unchanged.append(file_path)
        
        return changed, unchanged
    
    def clear_cache(self) -> bool:
        """
        Очищает кэш (удаляет файл кэша).
        
        Returns:
            bool: Успешность операции
        """
        try:
            if os.path.exists(self.cache_file):
                os.remove(self.cache_file)
                print(f"[КЭШ] Файл кэша удален: {self.cache_file}")
            
            self.cache_data = self._create_empty_cache()
            return True
            
        except Exception as e:
            print(f"[КЭШ] Ошибка при удалении кэша: {e}")
            return False
    
    def get_cache_info(self) -> dict:
        """
        Получает информацию о кэше.
        
        Returns:
            dict: Информация о кэше (количество файлов, дата последнего обновления и т.д.)
        """
        return {
            'program_name': self.cache_data['program_name'],
            'output_file': self.cache_data['output_file'],
            'file_count': len(self.cache_data['input_files']),
            'last_update': self.cache_data['last_update'],
            'cache_file_path': self.cache_file
        }
