import pandas
import requests
import time
from config import token
import datetime

"""
Парсер групп в Вконтакте (vk) по списку введенных групп
Данный парсер анализирует группы Вконтакте с целью выявления активных пользователей
Данная идея пришла при выборе группы для покупки рекламы,
ведь чем больше активной аудитории, тем выше будет охват рекламы в данной группе.

По всем возникшим вопросам, можете писать в группу https://vk.com/happython
"""


def get_offset(group_id):
    """Выявляем параметр offset для групп, 1 смещение * 1000 id"""
    params = {'access_token': token, 'group_id': group_id, 'v': 5.131}
    r = requests.get('https://api.vk.com/method/groups.getMembers', params=params)
    count = r.json()['response']['count']
    print(f'Количество подписчиков: {count}')
    if count > 1000:
        return count // 1000
    else:
        count = 1
        return count


def save_excel(data: list, filename: str):
    """сохранение результата в excel файл"""
    df = pandas.DataFrame(data)
    writer = pandas.ExcelWriter(f'{filename}.xlsx')
    df.to_excel(writer, sheet_name='data', index=False)
    # указываем размеры каждого столбца в итоговом файле
    writer.sheets['data'].set_column(0, 1, width=10)
    writer.sheets['data'].set_column(1, 2, width=12)
    writer.sheets['data'].set_column(2, 3, width=14)
    writer.sheets['data'].set_column(3, 4, width=23)
    writer.sheets['data'].set_column(4, 5, width=10)
    writer.sheets['data'].set_column(5, 6, width=8)
    writer.sheets['data'].set_column(6, 7, width=12)
    writer.close()
    print(f'Все сохранено в {filename}.xlsx\n')


def get_users(group_id, from_data):
    """Получаем всех участников группы и фильтруем от неактивных"""
    active_list = []
    users_can_closed_visit = []
    un_active_list = []
    data_list = []
    for offset in range(0, get_offset(group_id)+1):
        params = {'access_token': token,
                  'v': 5.131,
                  'group_id': group_id,
                  'offset': offset*1000,
                  'fields': 'last_seen, city, bdate, sex'}
        users = requests.get('https://api.vk.com/method/groups.getMembers', params=params).json()['response']
        for user in users['items']:
            user_id = user.get('id')
            if user.get('deactivated'):
                activity = user.get('deactivated')
            else:
                activity = 'active'
            try:
                last_seen = datetime.datetime.utcfromtimestamp(user.get('last_seen').get('time')).strftime('%d.%m.%Y')
            except AttributeError:
                last_seen = None
            if user.get('city'):
                city = user['city'].get('title')
            else:
                city = None
            bdate = user.get('bdate')
            if user.get('sex') == 2:
                sex = 'муж'
            else:
                sex = 'жен'
            try:
                platform = user.get('last_seen').get('platform')
            except AttributeError:
                platform = None
            if user.get('is_closed'):
                is_closed = 'закрытй профиль'
            else:
                is_closed = 'открытый профиль'
            data_list.append({
                'id': user_id,
                'activity': activity,
                'last_seen': last_seen,
                'city': city,
                'bdate': bdate,
                'sex': sex,
                'platform': platform,
                'is_closed': is_closed,
            })
            # проверка последнего посещения, не ранее указанной даты from_data преобразованной в timestamp
            start_point_data = datetime.datetime.strptime(from_data, '%d.%m.%Y').timestamp()
            try:
                if user['last_seen']['time'] >= start_point_data:
                    active_list.append(user['id'])
                else:
                    un_active_list.append(user['id'])
            except:
                users_can_closed_visit.append(user['id'])
    print(f"Активных подписчиков:   {len(active_list)} ({round((len(active_list)-len(un_active_list))/ (users['count']) * 100, 2)}%)")
    print(f'Не активные за отведенное время: {len(un_active_list)}\nЗабаненные/Удаленные: {len(users_can_closed_visit)}\n')
    return data_list


def parser(group_list):
    from_data = input('Введите дату, с которой хотите отслеживать активность\nв формате: дд.мм.гггг: ')
    # from_data = '20.08.2022'
    # all_active_users = []
    print(f'Анализируем с {from_data}\n')
    for group in group_list:
        print(f'Группа: {group}')
        # try:
        users = get_users(group_id=group, from_data=from_data)
        # all_active_users.extend(users)
        save_excel(users, filename=f"users_of_{group}")
        # time.sleep(2)
        # except Exception as ex:
        #     print(f'{group} - не предвиденная ошибка: {ex}\n')
        #     continue


if __name__ == '__main__':
    # вносим в список интересующие вас группы
    # group_list = ['happython', 'python_forum', 'vk_python', 'pirsipy']
    # group_list = ['happython']
    group_list = ['parsers_wildberries']
    parser(group_list)
