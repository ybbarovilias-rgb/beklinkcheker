import sys
import os
import json
import pickle
from datetime import datetime, timedelta
import pandas as pd
import requests
from bs4 import BeautifulSoup
import gzip
import random
import threading
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urlparse

from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *

class ModernStyle:
    """Класс для хранения стилей"""
    BG_COLOR = "#f0f0f0"
    ACCENT_COLOR = "#0078d4"
    TEXT_COLOR = "#333333"
    BORDER_COLOR = "#d0d0d0"
    HOVER_COLOR = "#e6f0fa"
    DISABLED_COLOR = "#cccccc"
    SUCCESS_COLOR = "#107c10"
    ERROR_COLOR = "#d13438"
    WARNING_COLOR = "#f2c335"

class ProjectManager:
    """Менеджер проектов и статистики"""
    
    def __init__(self):
        self.current_project = None
        self.projects_dir = "projects"
        self.intermediate_results = []  # Для хранения промежуточных результатов
        self.processed_count = 0  # Счетчик обработанных записей
        self.last_save_time = datetime.now()  # Время последнего сохранения
        if not os.path.exists(self.projects_dir):
            os.makedirs(self.projects_dir)
    
    def create_project(self, file_path):
        """Создание проекта для файла"""
        filename = os.path.basename(file_path)
        project_name = os.path.splitext(filename)[0]
        project_dir = os.path.join(self.projects_dir, project_name)
        
        if not os.path.exists(project_dir):
            os.makedirs(project_dir)
        
        # Копируем исходный файл в проектную папку
        project_file = os.path.join(project_dir, filename)
        if not os.path.exists(project_file):
            import shutil
            shutil.copy2(file_path, project_file)
        
        self.current_project = {
            'name': project_name,
            'dir': project_dir,
            'file': project_file,
            'stats_file': os.path.join(project_dir, 'stats.json'),
            'last_row': 0,
            'stats': {
                'dofollow': 0,
                'nofollow': 0,
                'text': 0,
                'errors': 0,
                'not_found': 0,
                'total_processed': 0,
                'last_processed': None
            }
        }
        
        self.intermediate_results = []
        self.processed_count = 0
        self.last_save_time = datetime.now()
        self.load_project_stats()
        return self.current_project
    
    def load_project_stats(self):
        """Загрузка статистики проекта"""
        if self.current_project and os.path.exists(self.current_project['stats_file']):
            try:
                with open(self.current_project['stats_file'], 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.current_project['last_row'] = data.get('last_row', 0)
                    self.current_project['stats'] = data.get('stats', self.current_project['stats'])
            except Exception as e:
                print(f"Ошибка загрузки статистики проекта: {e}")
    
    def save_project_stats(self, last_row=None):
        """Сохранение статистики проекта"""
        if self.current_project:
            if last_row is not None:
                self.current_project['last_row'] = last_row
            
            self.current_project['stats']['last_processed'] = datetime.now().isoformat()
            
            data = {
                'last_row': self.current_project['last_row'],
                'stats': self.current_project['stats']
            }
            
            try:
                with open(self.current_project['stats_file'], 'w', encoding='utf-8') as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
                self.last_save_time = datetime.now()  # Обновляем время последнего сохранения
            except Exception as e:
                print(f"Ошибка сохранения статистики проекта: {e}")
    
    def update_stats(self, result):
        """Обновление статистики с авто-сохранением каждые 10 ссылок"""
        if not self.current_project:
            return
            
        stats = self.current_project['stats']
        stats['total_processed'] += 1
        
        if result.get('status') == 'not_found':
            stats['not_found'] += 1
        elif result.get('link_type') == 'text':
            stats['text'] += 1
        elif result.get('follow_type') == 'dofollow':
            stats['dofollow'] += 1
        elif result.get('follow_type') == 'nofollow':
            stats['nofollow'] += 1
        elif result.get('status') == 'error':
            stats['errors'] += 1
        
        # Авто-сохранение каждые 10 ссылок
        if stats['total_processed'] % 10 == 0:
            self.save_project_stats()
            print(f"Статистика обновлена: {stats['total_processed']} ссылок обработано")
    
    def add_intermediate_result(self, result):
        """Добавление результата в промежуточные данные"""
        self.intermediate_results.append(result)
        self.processed_count += 1
        
        # Сохраняем промежуточные результаты каждые 100 записей
        if self.processed_count % 100 == 0:
            self.save_intermediate_results()
    
    def save_intermediate_results(self):
        """Сохранение промежуточных результатов"""
        if not self.current_project or not self.intermediate_results:
            return
            
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            intermediate_file = os.path.join(
                self.current_project['dir'], 
                f"intermediate_results_{timestamp}_{self.processed_count}.json"
            )
            
            with open(intermediate_file, 'w', encoding='utf-8') as f:
                json.dump(self.intermediate_results, f, ensure_ascii=False, indent=2)
            
            # Очищаем буфер после сохранения
            self.intermediate_results = []
            
            print(f"Промежуточные результаты сохранены: {intermediate_file}")
            
        except Exception as e:
            print(f"Ошибка сохранения промежуточных результатов: {e}")
    
    def save_final_results_and_cleanup(self):
        """Сохранение финальных результатов и очистка промежуточных файлов"""
        if not self.current_project:
            return
            
        # Сохраняем оставшиеся промежуточные результаты
        if self.intermediate_results:
            self.save_intermediate_results()
        
        # Собираем все промежуточные файлы и создаем общий отчет
        project_dir = self.current_project['dir']
        all_results = []
        
        # Собираем данные из всех промежуточных файлов
        for filename in os.listdir(project_dir):
            if filename.startswith('intermediate_results_') and filename.endswith('.json'):
                try:
                    with open(os.path.join(project_dir, filename), 'r', encoding='utf-8') as f:
                        results = json.load(f)
                        all_results.extend(results)
                except Exception as e:
                    print(f"Ошибка чтения промежуточного файла {filename}: {e}")
        
        # Создаем финальный отчет
        if all_results:
            self.create_final_report(all_results)
        
        # Удаляем промежуточные файлы
        self.cleanup_intermediate_files()
    
    def create_final_report(self, all_results):
        """Создание финального отчета"""
        if not self.current_project:
            return
            
        project_dir = self.current_project['dir']
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        
        # Разделение результатов
        dofollow_links = []
        nofollow_links = []
        text_links = []
        not_found = []
        
        for result in all_results:
            if result.get('status') == 'not_found':
                not_found.append(result)
            elif result.get('link_type') == 'text':
                text_links.append(result)
            elif result.get('follow_type') == 'dofollow':
                dofollow_links.append(result)
            elif result.get('follow_type') == 'nofollow':
                nofollow_links.append(result)
        
        # Сохранение результатов
        self.save_csv_to_project(dofollow_links, os.path.join(project_dir, f"dofollow_links_{timestamp}.csv"))
        self.save_csv_to_project(nofollow_links, os.path.join(project_dir, f"nofollow_links_{timestamp}.csv"))
        self.save_csv_to_project(text_links, os.path.join(project_dir, f"text_links_{timestamp}.csv"))
        self.save_csv_to_project(not_found, os.path.join(project_dir, f"not_found_{timestamp}.csv"))
        
        # Сохранение общего отчета
        self.save_csv_to_project(all_results, os.path.join(project_dir, f"full_report_{timestamp}.csv"))
        
        print(f"Финальный отчет сохранен в {project_dir}")
    
    def cleanup_intermediate_files(self):
        """Удаление промежуточных файлов"""
        if not self.current_project:
            return
            
        project_dir = self.current_project['dir']
        deleted_count = 0
        
        for filename in os.listdir(project_dir):
            if filename.startswith('intermediate_results_') and filename.endswith('.json'):
                try:
                    os.remove(os.path.join(project_dir, filename))
                    deleted_count += 1
                except Exception as e:
                    print(f"Ошибка удаления файла {filename}: {e}")
        
        print(f"Удалено промежуточных файлов: {deleted_count}")
    
    def save_csv_to_project(self, data, filename):
        """Сохранение данных в CSV файл"""
        if not data:
            return
            
        import csv
        with open(filename, 'w', newline='', encoding='utf-8-sig') as file:
            writer = csv.writer(file, delimiter=';')
            
            # Заголовки
            headers = ['Донорский URL', 'Найденная ссылка', 'Тип ссылки', 'Follow/Nofollow', 'Текст анкора']
            writer.writerow(headers)
            
            # Данные
            for item in data:
                row = [
                    item.get('donor_url', ''),
                    item.get('found_url', ''),
                    item.get('link_type', ''),
                    item.get('follow_type', ''),
                    item.get('anchor_text', '')
                ]
                writer.writerow(row)

class ProxyManager:
    """Менеджер прокси с сохранением в файл"""
    
    def __init__(self):
        self.proxies_file = "proxies.dat"
        self.config_file = "config.json"
        self.working_proxies = []
        self.api_key = ""
        self.last_check = None
        self.domains = []
        # API параметры
        self.perpage = 20
        self.country = "RU"
        self.country_not = ""
        
    def save_proxies(self):
        """Сохранение рабочих прокси в файл"""
        try:
            data = {
                'proxies': self.working_proxies,
                'timestamp': datetime.now().isoformat(),
                'api_key': self.api_key,
                'perpage': self.perpage,
                'country': self.country,
                'country_not': self.country_not
            }
            with open(self.proxies_file, 'wb') as f:
                pickle.dump(data, f)
        except Exception as e:
            print(f"Ошибка сохранения прокси: {e}")
    
    def load_proxies(self):
        """Загрузка рабочих прокси из файла"""
        try:
            if os.path.exists(self.proxies_file):
                with open(self.proxies_file, 'rb') as f:
                    data = pickle.load(f)
                    self.working_proxies = data.get('proxies', [])
                    self.api_key = data.get('api_key', '')
                    self.perpage = data.get('perpage', 20)
                    self.country = data.get('country', 'RU')
                    self.country_not = data.get('country_not', '')
                    timestamp_str = data.get('timestamp')
                    if timestamp_str:
                        self.last_check = datetime.fromisoformat(timestamp_str)
                    return True
        except Exception as e:
            print(f"Ошибка загрузки прокси: {e}")
        return False
    
    def save_config(self):
        """Сохранение конфигурации"""
        try:
            config = {
                'api_key': self.api_key,
                'perpage': self.perpage,
                'country': self.country,
                'country_not': self.country_not
            }
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"Ошибка сохранения конфигурации: {e}")
    
    def load_config(self):
        """Загрузка конфигурации"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    self.api_key = config.get('api_key', '')
                    self.perpage = config.get('perpage', 20)
                    self.country = config.get('country', 'RU')
                    self.country_not = config.get('country_not', '')
                    return True
        except Exception as e:
            print(f"Ошибка загрузки конфигурации: {e}")
        return False
    
    def add_working_proxy(self, proxy):
        """Добавление рабочего прокси"""
        if proxy not in self.working_proxies:
            self.working_proxies.append(proxy)
            self.save_proxies()
    
    def remove_proxy(self, proxy):
        """Удаление прокси"""
        if proxy in self.working_proxies:
            self.working_proxies.remove(proxy)
            self.save_proxies()
    
    def set_api_key(self, api_key):
        """Установка API ключа"""
        self.api_key = api_key
        self.save_config()
        self.save_proxies()  # Обновляем сохраненные данные

class LinkParser:
    def __init__(self, proxy_manager, project_manager):
        self.proxy_manager = proxy_manager
        self.project_manager = project_manager
        self.base_headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
            'Accept-Encoding': 'gzip, deflate',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        }
        self.stop_flag = False
        
    def get_session_with_proxy(self, proxy_string=None):
        """Создание сессии с прокси"""
        session = requests.Session()
        session.headers.update(self.base_headers)
        
        if proxy_string:
            try:
                # Определяем тип прокси
                if proxy_string.startswith('socks4://'):
                    proxy_type = 'socks4'
                    proxy_address = proxy_string[9:]
                elif proxy_string.startswith('socks5://'):
                    proxy_type = 'socks5'
                    proxy_address = proxy_string[9:]
                elif proxy_string.startswith('https://'):
                    proxy_type = 'https'
                    proxy_address = proxy_string[8:]
                elif proxy_string.startswith('http://'):
                    proxy_type = 'http'
                    proxy_address = proxy_string[7:]
                else:
                    proxy_type = 'http'
                    proxy_address = proxy_string
                
                proxies = {proxy_type: f"{proxy_type}://{proxy_address}"}
                session.proxies.update(proxies)
                
            except Exception as e:
                print(f"Ошибка настройки прокси {proxy_string}: {str(e)}")
        
        return session
    
    def fetch_page_with_proxy_retry(self, url, timeout=15):
        """Загрузка страницы с попытками через прокси"""
        # Первая попытка без прокси
        try:
            session = self.get_session_with_proxy()
            response = session.get(url, timeout=timeout, allow_redirects=True)
            response.raise_for_status()
            
            # Обработка различных сценариев декодирования
            content = self.decode_content(response)
            if content:
                return content
                
        except Exception as e:
            print(f"Ошибка без прокси {url}: {str(e)}")
        
        # Попытки через прокси
        if self.proxy_manager.working_proxies:
            # Выбираем 2 случайных прокси
            available_proxies = self.proxy_manager.working_proxies.copy()
            random.shuffle(available_proxies)
            
            for proxy in available_proxies[:2]:  # Максимум 2 попытки
                try:
                    session = self.get_session_with_proxy(proxy)
                    response = session.get(url, timeout=timeout, allow_redirects=True)
                    response.raise_for_status()
                    
                    content = self.decode_content(response)
                    if content:
                        print(f"Успешно через прокси: {url}")
                        return content
                        
                except Exception as e:
                    print(f"Ошибка через прокси {proxy} для {url}: {str(e)}")
                    continue
        
        return None
    
    def decode_content(self, response):
        """Улучшенная функция декодирования контента"""
        try:
            content = None
            
            # Если сервер отправил gzip, но он поврежден
            if response.headers.get('content-encoding') == 'gzip':
                try:
                    # Попробуем декодировать стандартным способом
                    content = response.text
                except Exception:
                    # Если не удалось, попробуем вручную
                    try:
                        content = gzip.decompress(response.content).decode('utf-8')
                    except Exception:
                        # Если и это не помогло, используем сырые данные
                        content = response.content.decode('utf-8', errors='ignore')
            else:
                content = response.text
                
            return content
            
        except Exception as e:
            print(f"Ошибка декодирования контента: {str(e)}")
            return None
    
    def fetch_page(self, url, timeout=15):
        """Улучшенная функция загрузки страницы с обработкой ошибок декодирования"""
        try:
            return self.fetch_page_with_proxy_retry(url, timeout)
        except Exception as e:
            print(f"Общая ошибка при загрузке {url}: {str(e)}")
            return None
    
    def check_link_follow_type(self, element):
        """Проверяет тип ссылки (dofollow/nofollow)"""
        if element.name != 'a':
            return 'text'
            
        rel = element.get('rel', [])
        if isinstance(rel, list):
            if 'nofollow' in rel:
                return 'nofollow'
        elif isinstance(rel, str):
            if 'nofollow' in rel:
                return 'nofollow'
        return 'dofollow'
    
    def find_target_url(self, soup, target_url):
        """Поиск целевого URL на странице"""
        if not target_url:
            return None
            
        links = soup.find_all('a', href=True)
        for link in links:
            href = link.get('href')
            if href and str(target_url).strip() in href:
                return {
                    'element': link,
                    'url': href,
                    'type': 'link',
                    'follow_type': self.check_link_follow_type(link),
                    'anchor_text': link.get_text(strip=True)
                }
        return None
    
    def find_anchor_text(self, soup, anchor_text):
        """Поиск анкора на странице"""
        if not anchor_text:
            return None
            
        # Поиск в ссылках
        links = soup.find_all('a')
        for link in links:
            link_text = link.get_text()
            if link_text and str(anchor_text).lower().strip() in link_text.lower():
                return {
                    'element': link,
                    'url': link.get('href', ''),
                    'type': 'link',
                    'follow_type': self.check_link_follow_type(link),
                    'anchor_text': link_text.strip()
                }
        
        # Поиск в тексте
        text_elements = soup.find_all(string=True)
        for element in text_elements:
            if element.strip() and str(anchor_text).lower().strip() in element.lower():
                return {
                    'element': element.parent,
                    'url': '',
                    'type': 'text',
                    'follow_type': 'text',
                    'anchor_text': element.strip()
                }
        return None
    
    def find_domain_links(self, soup, domains):
        """Поиск ссылок на указанные домены"""
        if not domains:
            return []
            
        links = soup.find_all('a', href=True)
        found_links = []
        
        for link in links:
            href = link.get('href')
            if not href:
                continue
                
            try:
                # Обработка относительных URL
                if href.startswith('//'):
                    href = 'http:' + href
                elif href.startswith('/'):
                    # Нужно получить базовый URL для относительных ссылок
                    pass
                elif not href.startswith(('http://', 'https://')):
                    continue
                    
                parsed_url = urlparse(href)
                domain = parsed_url.netloc.lower()
                
                # Проверяем каждый домен
                for target_domain in domains:
                    target_domain = target_domain.lower().strip()
                    if target_domain and target_domain in domain:
                        found_links.append({
                            'element': link,
                            'url': href,
                            'type': 'link',
                            'follow_type': self.check_link_follow_type(link),
                            'anchor_text': link.get_text(strip=True)
                        })
                        break
            except Exception as e:
                print(f"Ошибка при обработке ссылки {href}: {str(e)}")
                continue
                
        return found_links
    
    def parse_donor(self, donor_url, target_url=None, anchor_text=None):
        """Парсинг одного донорского URL"""
        if self.stop_flag:
            return {'donor_url': donor_url, 'status': 'stopped'}
        
        try:
            # Очистка URL от лишних пробелов
            donor_url = str(donor_url).strip()
            if not donor_url:
                return {'donor_url': donor_url, 'status': 'invalid_url'}
                
            # Добавляем протокол если его нет
            if not donor_url.startswith(('http://', 'https://')):
                donor_url = 'http://' + donor_url
            
            # Этап 1: поиск по целевому URL
            if target_url and str(target_url).strip():
                html = self.fetch_page(donor_url)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    result = self.find_target_url(soup, str(target_url).strip())
                    if result:
                        return {
                            'donor_url': donor_url,
                            'found_url': result['url'],
                            'link_type': result['type'],
                            'follow_type': result['follow_type'],
                            'anchor_text': result['anchor_text'],
                            'status': 'found_stage1'
                        }
            
            # Этап 2: поиск по анкору
            if anchor_text and str(anchor_text).strip():
                html = self.fetch_page(donor_url)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    result = self.find_anchor_text(soup, str(anchor_text).strip())
                    if result:
                        return {
                            'donor_url': donor_url,
                            'found_url': result['url'],
                            'link_type': result['type'],
                            'follow_type': result['follow_type'],
                            'anchor_text': result['anchor_text'],
                            'status': 'found_stage2'
                        }
            
            # Этап 3: поиск по доменам
            if self.proxy_manager.domains:
                html = self.fetch_page(donor_url)
                if html:
                    soup = BeautifulSoup(html, 'html.parser')
                    results = self.find_domain_links(soup, self.proxy_manager.domains)
                    if results:
                        # Возвращаем первую найденную ссылку
                        result = results[0]
                        return {
                            'donor_url': donor_url,
                            'found_url': result['url'],
                            'link_type': result['type'],
                            'follow_type': result['follow_type'],
                            'anchor_text': result['anchor_text'],
                            'status': 'found_stage3'
                        }
            
            # Не найдено
            return {
                'donor_url': donor_url,
                'found_url': '',
                'link_type': '',
                'follow_type': '',
                'anchor_text': '',
                'status': 'not_found'
            }
            
        except Exception as e:
            print(f"Ошибка при парсинге {donor_url}: {str(e)}")
            return {
                'donor_url': donor_url,
                'found_url': '',
                'link_type': '',
                'follow_type': '',
                'anchor_text': '',
                'status': 'error'
            }
    
    def parse_all(self, donor_urls, target_urls, anchors, domains, num_threads, progress_callback, start_row=0):
        """Парсинг всех URL с многопоточностью"""
        self.proxy_manager.domains = domains
        results = []
        total = len(donor_urls)
        completed = 0
        
        if total == 0:
            return results
            
        # Подготовка задач (начиная с указанной строки)
        tasks = []
        for i in range(start_row, total):
            donor_url = donor_urls[i]
            target_url = target_urls[i] if i < len(target_urls) else None
            anchor_text = anchors[i] if i < len(anchors) else None
            tasks.append((i, donor_url, target_url, anchor_text))
        
        # Многопоточный парсинг
        with ThreadPoolExecutor(max_workers=num_threads) as executor:
            future_to_task = {
                executor.submit(self.parse_donor, task[1], task[2], task[3]): task[0] 
                for task in tasks
            }
            
            for future in as_completed(future_to_task):
                if self.stop_flag:
                    break
                    
                try:
                    result = future.result()
                    row_index = future_to_task[future]
                    results.append(result)
                    completed += 1
                    
                    # Добавляем результат в промежуточные данные
                    self.project_manager.add_intermediate_result(result)
                    self.project_manager.update_stats(result)
                    
                    progress = (completed / len(tasks)) * 100 if tasks else 0
                    
                    # Повторно читаем stats.json каждые 5 секунд
                    current_time = datetime.now()
                    if (current_time - self.project_manager.last_save_time).seconds > 5:
                        self.project_manager.load_project_stats()
                        self.project_manager.last_save_time = current_time
                    
                    # Используем QTimer для безопасного обновления GUI
                    QTimer.singleShot(0, lambda: progress_callback(progress, result, row_index))
                        
                except Exception as e:
                    completed += 1
                    error_result = {'status': 'error'}
                    self.project_manager.add_intermediate_result(error_result)
                    self.project_manager.update_stats(error_result)
                    print(f"Ошибка при обработке задачи: {str(e)}")
                    progress = (completed / len(tasks)) * 100 if tasks else 0
                    
                    # Повторно читаем stats.json каждые 5 секунд
                    current_time = datetime.now()
                    if (current_time - self.project_manager.last_save_time).seconds > 5:
                        self.project_manager.load_project_stats()
                        self.project_manager.last_save_time = current_time
                    
                    QTimer.singleShot(0, lambda: progress_callback(progress, {'status': 'error'}, -1))
        
        return results

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.proxy_manager = ProxyManager()
        self.project_manager = ProjectManager()
        self.parser = LinkParser(self.proxy_manager, self.project_manager)
        self.current_stats = {
            'dofollow': 0,
            'nofollow': 0,
            'text': 0,
            'errors': 0,
            'not_found': 0,
            'total_processed': 0,
            'current_row': 0,
            'total_rows': 0
        }
        self.setup_ui()
        self.load_saved_data()
        
        # Обработчик закрытия приложения
        self.closeEvent = self.on_close
    
    def on_close(self, event):
        """Обработчик закрытия приложения"""
        self.save_final_results()
        event.accept()
    
    def setup_ui(self):
        self.setWindowTitle("Парсер ссылок - Qt5")
        self.setGeometry(100, 100, 1000, 700)
        
        # Центральный виджет
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        
        # Основной layout
        main_layout = QVBoxLayout(central_widget)
        
        # Создание вкладок
        tab_widget = QTabWidget()
        main_layout.addWidget(tab_widget)
        
        # Вкладка основных настроек
        self.main_tab = self.create_main_tab()
        tab_widget.addTab(self.main_tab, "Основные")
        
        # Вкладка прокси
        self.proxy_tab = self.create_proxy_tab()
        tab_widget.addTab(self.proxy_tab, "Прокси")
        
        # Вкладка статистики
        self.stats_tab = self.create_stats_tab()
        tab_widget.addTab(self.stats_tab, "Статистика")
        
        # Статус бар
        self.status_bar = QStatusBar()
        self.setStatusBar(self.status_bar)
        self.status_label = QLabel("Готов к работе")
        self.status_bar.addWidget(self.status_label)
        
    def create_main_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Файл
        file_group = QGroupBox("Файл данных")
        file_layout = QVBoxLayout(file_group)
        
        file_hlayout = QHBoxLayout()
        self.file_path_edit = QLineEdit()
        self.file_browse_btn = QPushButton("Обзор")
        self.file_browse_btn.clicked.connect(self.browse_file)
        file_hlayout.addWidget(self.file_path_edit)
        file_hlayout.addWidget(self.file_browse_btn)
        file_layout.addLayout(file_hlayout)
        
        layout.addWidget(file_group)
        
        # Информация о проекте
        self.project_info_group = QGroupBox("Информация о проекте")
        project_info_layout = QFormLayout(self.project_info_group)
        
        self.last_row_label = QLabel("0")
        self.last_processed_label = QLabel("Никогда")
        self.total_processed_label = QLabel("0")
        
        project_info_layout.addRow("Последняя обработанная строка:", self.last_row_label)
        project_info_layout.addRow("Последняя обработка:", self.last_processed_label)
        project_info_layout.addRow("Всего обработано:", self.total_processed_label)
        
        layout.addWidget(self.project_info_group)
        
        # Основная область - две колонки
        main_area_layout = QHBoxLayout()
        
        # Левая колонка - настройки обработки
        left_column = QVBoxLayout()
        
        # Настройки начала обработки
        start_group = QGroupBox("Настройки обработки")
        start_layout = QHBoxLayout(start_group)
        
        start_layout.addWidget(QLabel("Начать с строки:"))
        self.start_row_spin = QSpinBox()
        self.start_row_spin.setRange(0, 1000000)
        self.start_row_spin.setValue(0)
        start_layout.addWidget(self.start_row_spin)
        start_layout.addStretch()
        
        left_column.addWidget(start_group)
        
        # Колонки
        columns_group = QGroupBox("Колонки Excel")
        columns_layout = QFormLayout(columns_group)
        
        self.donor_combo = QComboBox()
        self.target_combo = QComboBox()
        self.anchor_combo = QComboBox()
        self.domains_edit = QLineEdit()
        
        columns_layout.addRow("Донорский URL*:", self.donor_combo)
        columns_layout.addRow("Искомый URL:", self.target_combo)
        columns_layout.addRow("Искомый анкор:", self.anchor_combo)
        columns_layout.addRow("Домены (через запятую):", self.domains_edit)
        
        left_column.addWidget(columns_group)
        
        # Настройки
        settings_group = QGroupBox("Настройки")
        settings_layout = QHBoxLayout(settings_group)
        
        settings_layout.addWidget(QLabel("Количество потоков:"))
        self.threads_spin = QSpinBox()
        self.threads_spin.setRange(1, 50)
        self.threads_spin.setValue(5)
        settings_layout.addWidget(self.threads_spin)
        settings_layout.addStretch()
        
        left_column.addWidget(settings_group)
        
        # Статистика текущей задачи
        stats_group = QGroupBox("Статистика текущей задачи")
        stats_layout = QGridLayout(stats_group)
        
        # Счетчики статистики
        self.dofollow_count_label = QLabel("0")
        self.dofollow_count_label.setStyleSheet(f"color: {ModernStyle.SUCCESS_COLOR}; font-weight: bold;")
        self.nofollow_count_label = QLabel("0")
        self.nofollow_count_label.setStyleSheet(f"color: {ModernStyle.ERROR_COLOR}; font-weight: bold;")
        self.text_count_label = QLabel("0")
        self.text_count_label.setStyleSheet(f"color: {ModernStyle.ACCENT_COLOR}; font-weight: bold;")
        self.errors_count_label = QLabel("0")
        self.errors_count_label.setStyleSheet(f"color: {ModernStyle.ERROR_COLOR}; font-weight: bold;")
        self.not_found_count_label = QLabel("0")
        self.not_found_count_label.setStyleSheet(f"color: {ModernStyle.WARNING_COLOR}; font-weight: bold;")
        self.processed_count_label = QLabel("0")
        self.processed_count_label.setStyleSheet(f"color: {ModernStyle.TEXT_COLOR}; font-weight: bold;")
        self.total_count_label = QLabel("0")
        self.total_count_label.setStyleSheet(f"color: {ModernStyle.TEXT_COLOR}; font-weight: bold;")
        
        stats_layout.addWidget(QLabel("Dofollow:"), 0, 0)
        stats_layout.addWidget(self.dofollow_count_label, 0, 1)
        stats_layout.addWidget(QLabel("Nofollow:"), 0, 2)
        stats_layout.addWidget(self.nofollow_count_label, 0, 3)
        stats_layout.addWidget(QLabel("Текст:"), 0, 4)
        stats_layout.addWidget(self.text_count_label, 0, 5)
        
        stats_layout.addWidget(QLabel("Ошибки:"), 1, 0)
        stats_layout.addWidget(self.errors_count_label, 1, 1)
        stats_layout.addWidget(QLabel("Не найдено:"), 1, 2)
        stats_layout.addWidget(self.not_found_count_label, 1, 3)
        stats_layout.addWidget(QLabel("Обработано:"), 1, 4)
        stats_layout.addWidget(self.processed_count_label, 1, 5)
        
        stats_layout.addWidget(QLabel("Всего:"), 2, 0)
        stats_layout.addWidget(self.total_count_label, 2, 1)
        
        left_column.addWidget(stats_group)
        
        # Кнопки управления
        controls_layout = QHBoxLayout()
        self.start_btn = QPushButton("Запустить")
        self.start_btn.clicked.connect(self.start_parsing)
        self.start_btn.setStyleSheet(f"background-color: {ModernStyle.ACCENT_COLOR}; color: white; padding: 8px;")
        
        self.continue_btn = QPushButton("Продолжить")
        self.continue_btn.clicked.connect(self.continue_parsing)
        self.continue_btn.setStyleSheet("padding: 8px;")
        
        self.stop_btn = QPushButton("Остановить")
        self.stop_btn.clicked.connect(self.stop_parsing)
        self.stop_btn.setEnabled(False)
        self.stop_btn.setStyleSheet("padding: 8px;")
        
        controls_layout.addWidget(self.start_btn)
        controls_layout.addWidget(self.continue_btn)
        controls_layout.addWidget(self.stop_btn)
        controls_layout.addStretch()
        
        left_column.addLayout(controls_layout)
        
        # Прогресс
        progress_layout = QVBoxLayout()
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        self.progress_bar.setFormat("%p%")  # Отображаем проценты
        progress_layout.addWidget(self.progress_bar)
        
        left_column.addLayout(progress_layout)
        left_column.addStretch()
        
        main_area_layout.addLayout(left_column)
        
        # Правая колонка - найденные домены
        right_column = QVBoxLayout()
        
        # Информация о найденных доменах (ТОЛЬКО из столбца "Искомый URL")
        self.domains_info_group = QGroupBox("Найденные домены (из столбца 'Искомый URL')")
        domains_info_layout = QVBoxLayout(self.domains_info_group)
        
        self.domains_info_label = QLabel("Загрузите файл и выберите столбец 'Искомый URL' для отображения доменов")
        self.domains_info_label.setWordWrap(True)
        self.domains_info_label.setStyleSheet("color: #666666; font-style: italic;")
        domains_info_layout.addWidget(self.domains_info_label)
        
        # Кнопка обновления списка доменов
        update_domains_btn = QPushButton("Обновить список доменов")
        update_domains_btn.clicked.connect(self.update_domains_list)
        domains_info_layout.addWidget(update_domains_btn)
        
        right_column.addWidget(self.domains_info_group)
        right_column.addStretch()
        
        main_area_layout.addLayout(right_column)
        
        layout.addLayout(main_area_layout)
        layout.addStretch()
        
        return widget
    
    def create_proxy_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # API ключ
        api_group = QGroupBox("API ключ")
        api_layout = QVBoxLayout(api_group)
        
        api_hlayout = QHBoxLayout()
        self.api_key_edit = QLineEdit()
        self.api_key_edit.setEchoMode(QLineEdit.Password)
        self.register_link = QLabel('<a href="#">Регистрация</a>')
        self.register_link.linkActivated.connect(self.open_registration)
        api_hlayout.addWidget(self.api_key_edit)
        api_hlayout.addWidget(self.register_link)
        api_layout.addLayout(api_hlayout)
        
        layout.addWidget(api_group)
        
        # API параметры
        api_params_group = QGroupBox("Параметры API")
        api_params_layout = QVBoxLayout(api_params_group)
        
        # Количество прокси (галочки с шагом 20 от 20 до 100)
        perpage_layout = QHBoxLayout()
        perpage_layout.addWidget(QLabel("Количество прокси:"))
        
        # Создаем радио-кнопки для выбора количества прокси
        self.perpage_group = QButtonGroup()
        self.perpage_buttons = []
        
        perpage_values = [20, 40, 60, 80, 100]
        for i, value in enumerate(perpage_values):
            radio_btn = QRadioButton(str(value))
            radio_btn.value = value
            self.perpage_group.addButton(radio_btn, i)
            self.perpage_buttons.append(radio_btn)
            perpage_layout.addWidget(radio_btn)
            
            # Устанавливаем значение по умолчанию (20)
            if value == 20:
                radio_btn.setChecked(True)
        
        api_params_layout.addLayout(perpage_layout)
        
        # Страна
        country_layout = QHBoxLayout()
        country_layout.addWidget(QLabel("Страна (пример: RU):"))
        self.country_edit = QLineEdit()
        self.country_edit.setText("RU")
        country_layout.addWidget(self.country_edit)
        api_params_layout.addLayout(country_layout)
        
        # Исключенные страны
        country_not_layout = QHBoxLayout()
        country_not_layout.addWidget(QLabel("Исключить страны (пример: RU,UA):"))
        self.country_not_edit = QLineEdit()
        country_not_layout.addWidget(self.country_not_edit)
        api_params_layout.addLayout(country_not_layout)
        
        layout.addWidget(api_params_group)
        
        # Управление прокси
        proxy_controls_group = QGroupBox("Управление прокси")
        proxy_controls_layout = QHBoxLayout(proxy_controls_group)
        
        self.get_proxy_btn = QPushButton("Получить прокси")
        self.get_proxy_btn.clicked.connect(self.get_proxies)
        self.check_proxy_btn = QPushButton("Проверить прокси")
        self.check_proxy_btn.clicked.connect(self.check_proxies)
        self.save_proxy_btn = QPushButton("Сохранить прокси")
        self.save_proxy_btn.clicked.connect(self.save_proxies)
        
        proxy_controls_layout.addWidget(self.get_proxy_btn)
        proxy_controls_layout.addWidget(self.check_proxy_btn)
        proxy_controls_layout.addWidget(self.save_proxy_btn)
        proxy_controls_layout.addStretch()
        
        layout.addWidget(proxy_controls_group)
        
        # Список прокси
        self.proxy_list = QListWidget()
        layout.addWidget(QLabel("Рабочие прокси:"))
        layout.addWidget(self.proxy_list)
        
        # Статус прокси
        self.proxy_status_label = QLabel("Прокси: 0")
        layout.addWidget(self.proxy_status_label)
        
        layout.addStretch()
        
        return widget
    
    def create_stats_tab(self):
        widget = QWidget()
        layout = QVBoxLayout(widget)
        
        # Статистика в виде сетки
        stats_grid = QGridLayout()
        
        # Счетчики
        self.dofollow_label = self.create_stat_label("Dofollow:", "0", ModernStyle.SUCCESS_COLOR)
        self.nofollow_label = self.create_stat_label("Nofollow:", "0", ModernStyle.ERROR_COLOR)
        self.text_label = self.create_stat_label("Текст:", "0", ModernStyle.ACCENT_COLOR)
        self.errors_label = self.create_stat_label("Ошибки:", "0", ModernStyle.ERROR_COLOR)
        self.not_found_label = self.create_stat_label("Не найдено:", "0", ModernStyle.WARNING_COLOR)
        self.processed_label = self.create_stat_label("Обработано:", "0", ModernStyle.TEXT_COLOR)
        
        stats_grid.addWidget(self.dofollow_label[0], 0, 0)
        stats_grid.addWidget(self.dofollow_label[1], 0, 1)
        stats_grid.addWidget(self.nofollow_label[0], 0, 2)
        stats_grid.addWidget(self.nofollow_label[1], 0, 3)
        stats_grid.addWidget(self.text_label[0], 0, 4)
        stats_grid.addWidget(self.text_label[1], 0, 5)
        
        stats_grid.addWidget(self.errors_label[0], 1, 0)
        stats_grid.addWidget(self.errors_label[1], 1, 1)
        stats_grid.addWidget(self.not_found_label[0], 1, 2)
        stats_grid.addWidget(self.not_found_label[1], 1, 3)
        stats_grid.addWidget(self.processed_label[0], 1, 4)
        stats_grid.addWidget(self.processed_label[1], 1, 5)
        
        layout.addLayout(stats_grid)
        layout.addStretch()
        
        return widget
    
    def create_stat_label(self, text, value, color):
        """Создание стилизованной метки статистики"""
        label_text = QLabel(text)
        label_value = QLabel(value)
        label_value.setStyleSheet(f"color: {color}; font-weight: bold;")
        label_value.setAlignment(Qt.AlignLeft)
        return (label_text, label_value)
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Выберите Excel файл", "", "Excel Files (*.xlsx *.xls)"
        )
        if file_path:
            self.file_path_edit.setText(file_path)
            self.load_columns()
            self.load_project_info()
            self.update_domains_list()
    
    def load_columns(self):
        try:
            file_path = self.file_path_edit.text()
            if not file_path:
                return
                
            df = pd.read_excel(file_path)
            columns = list(df.columns)
            
            self.donor_combo.clear()
            self.target_combo.clear()
            self.anchor_combo.clear()
            
            self.donor_combo.addItems(columns)
            self.target_combo.addItems([''] + columns)
            self.anchor_combo.addItems([''] + columns)
            
        except Exception as e:
            QMessageBox.critical(self, "Ошибка", f"Не удалось загрузить файл: {str(e)}")
    
    def load_project_info(self):
        """Загрузка информации о проекте"""
        file_path = self.file_path_edit.text()
        if not file_path:
            return
            
        # Создаем проект
        self.project_manager.create_project(file_path)
        project = self.project_manager.current_project
        
        if project:
            # Обновляем информацию в интерфейсе
            self.last_row_label.setText(str(project['last_row']))
            self.start_row_spin.setValue(project['last_row'])
            
            # Обновляем статистику
            stats = project['stats']
            self.total_processed_label.setText(str(stats['total_processed']))
            
            if stats['last_processed']:
                try:
                    last_processed = datetime.fromisoformat(stats['last_processed'])
                    self.last_processed_label.setText(last_processed.strftime("%d.%m.%Y %H:%M:%S"))
                except:
                    self.last_processed_label.setText("Ошибка даты")
            else:
                self.last_processed_label.setText("Никогда")
            
            # Обновляем статистику в табе статистики
            self.update_statistics_display_from_project()
    
    def update_statistics_display_from_project(self):
        """Обновление отображения статистики из проекта"""
        if self.project_manager.current_project:
            stats = self.project_manager.current_project['stats']
            self.dofollow_label[1].setText(str(stats['dofollow']))
            self.nofollow_label[1].setText(str(stats['nofollow']))
            self.text_label[1].setText(str(stats['text']))
            self.errors_label[1].setText(str(stats['errors']))
            self.not_found_label[1].setText(str(stats['not_found']))
            self.processed_label[1].setText(str(stats['total_processed']))
            self.total_processed_label.setText(str(stats['total_processed']))
    
    def load_saved_data(self):
        """Загрузка сохраненных данных при запуске"""
        # Загрузка конфигурации
        self.proxy_manager.load_config()
        self.api_key_edit.setText(self.proxy_manager.api_key)
        
        # Устанавливаем сохраненное значение perpage
        perpage_values = [20, 40, 60, 80, 100]
        if self.proxy_manager.perpage in perpage_values:
            index = perpage_values.index(self.proxy_manager.perpage)
            if index < len(self.perpage_buttons):
                self.perpage_buttons[index].setChecked(True)
        
        self.country_edit.setText(self.proxy_manager.country)
        self.country_not_edit.setText(self.proxy_manager.country_not)
        
        # Загрузка прокси
        if self.proxy_manager.load_proxies():
            self.update_proxy_list()
            # Автоматическая проверка прокси, если прошло больше 1 часа
            if self.proxy_manager.last_check:
                time_diff = datetime.now() - self.proxy_manager.last_check
                if time_diff > timedelta(hours=1):
                    self.check_proxies()
        else:
            self.status_label.setText("Нет сохраненных прокси")
    
    def update_proxy_list(self):
        """Обновление списка прокси в интерфейсе"""
        self.proxy_list.clear()
        for proxy in self.proxy_manager.working_proxies:
            self.proxy_list.addItem(proxy)
        self.proxy_status_label.setText(f"Прокси: {len(self.proxy_manager.working_proxies)}")
    
    def get_perpage_value(self):
        """Получение выбранного значения количества прокси"""
        checked_button = self.perpage_group.checkedButton()
        if checked_button:
            return checked_button.value
        return 20  # значение по умолчанию
    
    def get_proxies(self):
        """Получение списка прокси через API"""
        api_key = self.api_key_edit.text()
        if not api_key:
            QMessageBox.critical(self, "Ошибка", "Введите API ключ")
            return
            
        try:
            self.status_label.setText("Получение списка прокси...")
            QApplication.processEvents()
            
            # Сохраняем параметры API
            self.proxy_manager.perpage = self.get_perpage_value()
            self.proxy_manager.country = self.country_edit.text()
            self.proxy_manager.country_not = self.country_not_edit.text()
            
            # Формируем URL с параметрами
            api_url = f"https://htmlweb.ru/json/proxy/get?short=2&perpage={self.proxy_manager.perpage}&api_key={api_key}"
            
            # Добавляем параметр страны если указан
            if self.proxy_manager.country.strip():
                api_url += f"&country={self.proxy_manager.country}"
            
            # Добавляем параметр исключенных стран если указаны
            if self.proxy_manager.country_not.strip():
                api_url += f"&country_not={self.proxy_manager.country_not}"
            
            response = requests.get(api_url, timeout=30)
            response.raise_for_status()
            
            data = response.json()
            
            # Извлекаем прокси из ответа
            new_proxies = []
            for key, value in data.items():
                if key != 'limit' and isinstance(value, str):
                    new_proxies.append(value)
            
            # Добавляем новые прокси к существующим
            for proxy in new_proxies:
                if proxy not in self.proxy_manager.working_proxies:
                    self.proxy_manager.working_proxies.append(proxy)
            
            self.proxy_manager.set_api_key(api_key)
            self.update_proxy_list()
            self.status_label.setText("Прокси получены")
            QMessageBox.information(self, "Успех", f"Получено {len(new_proxies)} прокси")
            
        except Exception as e:
            self.status_label.setText("Ошибка получения прокси")
            QMessageBox.critical(self, "Ошибка", f"Не удалось получить прокси: {str(e)}")
    
    def check_proxies(self):
        """Проверка работоспособности прокси"""
        if not self.proxy_manager.working_proxies:
            QMessageBox.critical(self, "Ошибка", "Нет прокси для проверки")
            return
            
        self.status_label.setText("Проверка прокси...")
        QApplication.processEvents()
        
        # Проверяем несколько прокси параллельно
        working_proxies = []
        
        def check_single_proxy(proxy_string):
            try:
                # Парсим прокси
                if proxy_string.startswith('socks4://'):
                    proxy_type = 'socks4'
                    proxy_address = proxy_string[9:]
                elif proxy_string.startswith('socks5://'):
                    proxy_type = 'socks5'
                    proxy_address = proxy_string[9:]
                elif proxy_string.startswith('https://'):
                    proxy_type = 'https'
                    proxy_address = proxy_string[8:]
                elif proxy_string.startswith('http://'):
                    proxy_type = 'http'
                    proxy_address = proxy_string[7:]
                else:
                    proxy_type = 'http'
                    proxy_address = proxy_string
                
                proxies = {proxy_type: f"{proxy_type}://{proxy_address}"}
                
                # Проверяем на простом запросе
                test_session = requests.Session()
                test_session.headers.update({
                    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'
                })
                
                response = test_session.get('http://httpbin.org/ip', 
                                          proxies=proxies, 
                                          timeout=10)
                
                if response.status_code == 200:
                    return proxy_string
                return None
                
            except Exception:
                return None
        
        # Многопоточная проверка
        with ThreadPoolExecutor(max_workers=10) as executor:
            future_to_proxy = {executor.submit(check_single_proxy, proxy): proxy 
                             for proxy in self.proxy_manager.working_proxies[:20]}
            
            completed = 0
            total = min(20, len(self.proxy_manager.working_proxies))
            
            for future in as_completed(future_to_proxy):
                result = future.result()
                if result:
                    working_proxies.append(result)
                
                completed += 1
                self.status_label.setText(f"Проверка прокси... {completed}/{total}")
                QApplication.processEvents()
        
        # Обновляем список рабочих прокси
        self.proxy_manager.working_proxies = working_proxies
        self.proxy_manager.save_proxies()
        self.update_proxy_list()
        self.status_label.setText("Проверка прокси завершена")
        QMessageBox.information(self, "Успех", f"Работающих прокси: {len(working_proxies)}")
    
    def save_proxies(self):
        """Сохранение прокси в файл"""
        # Сохраняем текущие параметры API
        self.proxy_manager.perpage = self.get_perpage_value()
        self.proxy_manager.country = self.country_edit.text()
        self.proxy_manager.country_not = self.country_not_edit.text()
        
        self.proxy_manager.save_proxies()
        self.status_label.setText("Прокси сохранены")
    
    def start_parsing(self):
        """Запуск парсинга с нуля (обнуляет статистику текущей задачи)"""
        if not self.validate_inputs():
            return
            
        # Обнуляем статистику текущей задачи
        self.reset_current_statistics()
        
        self.parser.stop_flag = False
        self.start_btn.setEnabled(False)
        self.continue_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        
        # Запуск парсинга в отдельном потоке
        self.parsing_thread = threading.Thread(target=self.run_parsing, daemon=True)
        self.parsing_thread.start()
    
    def continue_parsing(self):
        """Продолжение парсинга с последней позиции (не обнуляет статистику текущей задачи)"""
        if not self.validate_inputs():
            return
            
        # Не обнуляем статистику текущей задачи, продолжаем с последней позиции
        self.parser.stop_flag = False
        self.start_btn.setEnabled(False)
        self.continue_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        
        # Устанавливаем начальную строку на последнюю обработанную
        if self.project_manager.current_project:
            last_row = self.project_manager.current_project['last_row']
            self.start_row_spin.setValue(last_row)
        
        # Запуск парсинга в отдельном потоке
        self.parsing_thread = threading.Thread(target=self.run_parsing, daemon=True)
        self.parsing_thread.start()
    
    def stop_parsing(self):
        """Остановка парсинга с сохранением результатов"""
        self.parser.stop_flag = True
        self.status_label.setText("Остановка и сохранение результатов...")
        self.save_final_results()
        self.status_label.setText("Парсинг остановлен")
        self.start_btn.setEnabled(True)
        self.continue_btn.setEnabled(True)
        self.stop_btn.setEnabled(False)
    
    def reset_current_statistics(self):
        """Сброс статистики текущей задачи"""
        self.current_stats = {
            'dofollow': 0,
            'nofollow': 0,
            'text': 0,
            'errors': 0,
            'not_found': 0,
            'total_processed': 0,
            'current_row': 0,
            'total_rows': 0
        }
        self.update_current_statistics_display()
    
    def update_current_statistics_display(self):
        """Обновление отображения статистики текущей задачи"""
        self.dofollow_count_label.setText(str(self.current_stats['dofollow']))
        self.nofollow_count_label.setText(str(self.current_stats['nofollow']))
        self.text_count_label.setText(str(self.current_stats['text']))
        self.errors_count_label.setText(str(self.current_stats['errors']))
        self.not_found_count_label.setText(str(self.current_stats['not_found']))
        self.processed_count_label.setText(str(self.current_stats['total_processed']))
        self.total_count_label.setText(str(self.current_stats['total_rows']))
    
    def save_final_results(self):
        """Сохранение финальных результатов"""
        try:
            self.project_manager.save_final_results_and_cleanup()
            self.status_label.setText("Результаты сохранены")
        except Exception as e:
            print(f"Ошибка при сохранении финальных результатов: {e}")
            self.status_label.setText("Ошибка сохранения результатов")
    
    def validate_inputs(self):
        if not self.file_path_edit.text():
            QMessageBox.critical(self, "Ошибка", "Выберите Excel файл")
            return False
            
        if not self.donor_combo.currentText():
            QMessageBox.critical(self, "Ошибка", "Выберите колонку с донорскими URL")
            return False
            
        return True
    
    def update_statistics(self, result):
        """Обновление статистики"""
        # Обновляем статистику текущей задачи
        self.current_stats['total_processed'] += 1
        
        if result.get('status') == 'not_found':
            self.current_stats['not_found'] += 1
        elif result.get('link_type') == 'text':
            self.current_stats['text'] += 1
        elif result.get('follow_type') == 'dofollow':
            self.current_stats['dofollow'] += 1
        elif result.get('follow_type') == 'nofollow':
            self.current_stats['nofollow'] += 1
        elif result.get('status') == 'error':
            self.current_stats['errors'] += 1
            
        self.update_current_statistics_display()
        
        # Обновляем статистику проекта
        self.project_manager.update_stats(result)
        self.update_statistics_display_from_project()
    
    def progress_callback(self, value, result, row_index):
        """Callback для обновления прогресса"""
        # Используем QTimer для безопасного обновления GUI
        QTimer.singleShot(0, lambda: self._safe_progress_update(value, result, row_index))
    
    def _safe_progress_update(self, value, result, row_index):
        """Безопасное обновление прогресса в основном потоке"""
        try:
            self.progress_bar.setValue(int(value))
            
            # Обновляем последнюю обработанную строку
            if row_index >= 0 and self.project_manager.current_project:
                self.project_manager.save_project_stats(row_index + 1)
                self.last_row_label.setText(str(row_index + 1))
                self.current_stats['current_row'] = row_index + 1
            
            self.update_statistics(result)
        except Exception as e:
            print(f"Ошибка обновления GUI: {e}")
    
    def run_parsing(self):
        try:
            # Чтение Excel файла из проекта
            project = self.project_manager.current_project
            df = pd.read_excel(project['file'])
            
            # Получение данных
            donor_urls = df[self.donor_combo.currentText()].dropna().tolist()
            target_urls = []
            anchors = []
            
            if self.target_combo.currentText():
                target_urls = df[self.target_combo.currentText()].dropna().tolist()
            if self.anchor_combo.currentText():
                anchors = df[self.anchor_combo.currentText()].dropna().tolist()
                
            domains = [d.strip() for d in self.domains_edit.text().split(',') if d.strip()]
            
            # Получаем начальную строку
            start_row = self.start_row_spin.value()
            
            # Устанавливаем общее количество строк для статистики
            self.current_stats['total_rows'] = len(donor_urls)
            self.update_current_statistics_display()
            
            # Запуск парсинга
            results = self.parser.parse_all(
                donor_urls, target_urls, anchors, domains,
                self.threads_spin.value(), self.progress_callback, start_row
            )
            
            # Сохранение финальных результатов
            self.save_final_results()
            
            if not self.parser.stop_flag:
                self.status_label.setText("Парсинг завершен!")
                QMessageBox.information(self, "Успех", "Парсинг завершен успешно!")
                
        except Exception as e:
            self.status_label.setText("Ошибка")
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка: {str(e)}")
        finally:
            self.start_btn.setEnabled(True)
            self.continue_btn.setEnabled(True)
            self.stop_btn.setEnabled(False)
    
    def open_registration(self):
        """Открытие страницы регистрации"""
        QDesktopServices.openUrl(QUrl("https://htmlweb.ru/user/signup.php"))
    
    def update_domains_list(self):
        """Обновление списка доменов из столбца 'Искомый URL'"""
        try:
            file_path = self.file_path_edit.text()
            if not file_path or not self.target_combo.currentText():
                self.domains_info_label.setText("Выберите файл и столбец 'Искомый URL' для отображения доменов")
                self.domains_info_label.setStyleSheet("color: #666666; font-style: italic;")
                return
                
            df = pd.read_excel(file_path)
            
            # Извлекаем домены ТОЛЬКО из столбца "Искомый URL"
            if self.target_combo.currentText():
                target_urls = df[self.target_combo.currentText()].dropna().tolist()
                domains = set()
                
                for url in target_urls:
                    try:
                        parsed_url = urlparse(str(url).strip())
                        domain = parsed_url.netloc.lower()
                        if domain:
                            domains.add(domain)
                    except Exception as e:
                        print(f"Ошибка парсинга URL {url}: {e}")
                        continue
                
                if domains:
                    domains_text = "\n".join(sorted(list(domains)))
                    self.domains_info_label.setText(f"Найдено доменов: {len(domains)}\n\n{domains_text}")
                    self.domains_info_label.setStyleSheet("color: #333333; font-style: normal;")
                else:
                    self.domains_info_label.setText("Домены не найдены в выбранном столбце")
                    self.domains_info_label.setStyleSheet("color: #666666; font-style: italic;")
            else:
                # Если столбец не выбран, очищаем список доменов
                self.domains_info_label.setText("Выберите столбец 'Искомый URL' для отображения доменов")
                self.domains_info_label.setStyleSheet("color: #666666; font-style: italic;")
                
        except Exception as e:
            print(f"Ошибка извлечения доменов: {e}")
            self.domains_info_label.setText("Ошибка извлечения доменов")
            self.domains_info_label.setStyleSheet("color: #d13438; font-style: italic;")

def main():
    app = QApplication(sys.argv)
    app.setStyle('Fusion')  # Современный стиль
    
    # Установка стилей
    stylesheet = f"""
    QMainWindow {{
        background-color: {ModernStyle.BG_COLOR};
    }}
    QGroupBox {{
        font-weight: bold;
        border: 1px solid {ModernStyle.BORDER_COLOR};
        border-radius: 5px;
        margin-top: 1ex;
        padding-top: 10px;
    }}
    QGroupBox::title {{
        subcontrol-origin: margin;
        left: 10px;
        padding: 0 5px 0 5px;
    }}
    QPushButton {{
        border: 1px solid {ModernStyle.BORDER_COLOR};
        border-radius: 4px;
        padding: 5px 15px;
        background-color: white;
    }}
    QPushButton:hover {{
        background-color: {ModernStyle.HOVER_COLOR};
    }}
    QPushButton:disabled {{
        background-color: {ModernStyle.DISABLED_COLOR};
        color: white;
    }}
    QProgressBar {{
        border: 1px solid {ModernStyle.BORDER_COLOR};
        border-radius: 3px;
        text-align: center;
    }}
    QProgressBar::chunk {{
        background-color: {ModernStyle.ACCENT_COLOR};
    }}
    QRadioButton {{
        margin-right: 10px;
    }}
    """
    
    app.setStyleSheet(stylesheet)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()