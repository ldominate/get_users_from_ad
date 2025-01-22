import pandas as pd
from ldap3 import Server, Connection, ALL, NTLM
from datetime import datetime
import argparse
import configparser

# Метод подключения и создания соединения с AD
def create_ad_connection(ad_server, ad_user, ad_password):
    try:
        server = Server(ad_server, get_info=ALL)
        conn = Connection(server, user=ad_user, password=ad_password, authentication=NTLM, auto_bind=True)
        print("[INFO] Подключение к Active Directory успешно.")
        return conn
    except Exception as e:
        print(f"[ERROR] Ошибка при подключении к AD: {e}")
        return None

# Проверяем права доступа указанного пользователя
def check_user_access(conn, search_filter):
    try:
        test_attribute = 'cn'
        print("[INFO] Проверка прав доступа...")

        if not conn.search(AD_SEARCH_BASE, search_filter, attributes=[test_attribute], size_limit=1):
            print("[ERROR] У пользователя недостаточно прав для выполнения поиска в AD.")
            print(f"LDAP Result: {conn.result}")
            return False

        print("[SUCCESS] Проверка прав доступа пройдена. Права на чтение объектов есть.")
        return True
    except Exception as e:
        print(f"[ERROR] Ошибка при проверке доступа: {e}")
        return False

# Получаем пользователей из AD
def get_ad_users(conn, search_filter):
    try:
        attributes = ['sAMAccountName', 'displayName', 'mail', 'title', 'department', 'userAccountControl', 'lastLogon']

        conn.search(AD_SEARCH_BASE, search_filter, attributes=attributes)

        if not conn.entries:
            print("[INFO] Результаты поиска пусты. Проверьте фильтр или настройки.")
            return []

        # Обработка найденных пользователей
        users = []
        for entry in conn.entries:
            user = entry.entry_attributes_as_dict
            def safe_get(attr, default=''):
                value = user.get(attr, [default])
                return value[0] if value else default
            
            users.append({
                'SamAccountName': safe_get('sAMAccountName'),
                'DisplayName': safe_get('displayName'),
                'Email': safe_get('mail'),
                'Title': safe_get('title'),
                'Department': safe_get('department'),
                'Enabled': not (int(safe_get('userAccountControl', [0])) & 2) if 'userAccountControl' in user else False,
                'LastLogon': convert_last_logon(safe_get('lastLogon', ['0'])) if 'lastLogon' in user else 'Never',
            })

        print(f"[INFO] Найдено {len(users)} пользователей.")
        return users

    except Exception as e:
        print(f"[ERROR] Ошибка при получении пользователей: {e}")
        return []

# Конвертация lastLogon
def convert_last_logon(last_logon):
    try:
        if last_logon and last_logon != '0':
            timestamp = int(last_logon) / 10000000 - 11644473600
            return datetime.fromtimestamp(timestamp).strftime('%Y-%m-%d %H:%M:%S')
        return 'Never'
    except:
        return 'Never'

# Сохранение данных в Excel
def save_to_excel(users, file_path):
    try:
        df = pd.DataFrame(users)
        df.to_excel(file_path, index=False, engine='openpyxl')
        print(f"[INFO] Данные сохранены в файл: {file_path}")
    except Exception as e:
        print(f"[ERROR] Ошибка сохранения в Excel: {e}")

# Основная функция
if __name__ == '__main__':

    parser = argparse.ArgumentParser(description="Скрипт для получения пользователей из Active Directory.")
    parser.add_argument('--config', type=str, help="Путь к файлу конфигурации.")
    parser.add_argument('--server', type=str, help="Адрес сервера AD.")
    parser.add_argument('--user', type=str, help="Учетная запись для подключения к AD.")
    parser.add_argument('--password', type=str, help="Пароль для подключения к AD.")
    parser.add_argument('--base', type=str, help="База поиска в AD.")
    parser.add_argument('--filter', type=str, default='(objectClass=user)', help="Фильтр поиска в AD.")
    parser.add_argument('--output', type=str, default='AD_Users.xlsx', help="Путь к выходному файлу Excel.")

    args = parser.parse_args()

    # Если параметры не заданы, выводим справку
    if not (args.config or (args.server and args.user and args.password and args.base)):
        parser.print_help()
        exit(0)

        # Чтение конфигурации из файла
    config = configparser.ConfigParser(interpolation=None)
    if args.config:
        config.read(args.config)
    
    AD_SERVER = args.server or config.get('AD', 'server', fallback=None)
    AD_USER = args.user or config.get('AD', 'user', fallback=None)
    AD_PASSWORD = args.password or config.get('AD', 'password', fallback=None)
    AD_SEARCH_BASE = args.base or config.get('AD', 'base', fallback=None)
    SEARCH_FILTER = args.filter
    output_file = args.output
    
    if not all([AD_SERVER, AD_USER, AD_PASSWORD, AD_SEARCH_BASE]):
        print("[ERROR] Не заданы все параметры подключения к AD.")
        exit(1)

    # Создание соединения с AD
    conn = create_ad_connection(AD_SERVER, AD_USER, AD_PASSWORD)

    if conn:
        # Проверка прав доступа
        if check_user_access(conn, SEARCH_FILTER):
            ad_users = get_ad_users(conn, SEARCH_FILTER)
            if ad_users:
                save_to_excel(ad_users, output_file)
            else:
                print("[INFO] Нет данных для сохранения.")
        else:
            print("[ERROR] Скрипт остановлен из-за недостаточных прав доступа.")

        conn.unbind()
    else:
        print("[ERROR] Не удалось установить соединение с Active Directory.")