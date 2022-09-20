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


def get_users(group_id, from_data):
    """Получаем всех участников группы и фильтруем от неактивных"""
    active_list = []
    users_can_closed_visit = []
    un_active_list = []
    for offset in range(0, get_offset(group_id)+1):
        params = {'access_token': token, 'v': 5.131, 'group_id': group_id, 'offset': offset*1000, 'fields': 'last_seen'}
        users = requests.get('https://api.vk.com/method/groups.getMembers', params=params).json()['response']
        for user in users['items']:
            # проверка последнего посещения, не ранее указанной даты from_data преобразованной в timestamp
            start_point_data = datetime.datetime.strptime(from_data, '%d.%m.%Y').timestamp()
            try:
                if user['last_seen']['time'] >= start_point_data:
                    active_list.append(user['id'])
                else:
                    un_active_list.append(user['id'])
            except:
                users_can_closed_visit.append(user['id'])
    print(f'Количество пользователей со скрытой датой: {len(users_can_closed_visit)}')
    print(f"Активных подписчиков:   {len(active_list)} ({round(len(active_list) / (users['count'] - len(un_active_list)) *100, 2)}%)")
    print(f'Не активные подписчики: {len(un_active_list)}\n')
    return active_list


def parser(group_list):
    from_data = input('Введите дату, с которой хотите отслеживать активность\nв формате: дд.мм.гггг: ')
    # from_data = '20.08.2022'
    all_active_users = []
    print(f'Анализируем с {from_data}\n')
    for group in group_list:
        print(f'Группа: {group}')
        try:
            users = get_users(group, from_data=from_data)
            all_active_users.extend(users)
            time.sleep(2)
        except Exception as ex:
            print(f'{group} - не предвиденная ошибка: {ex}\n')
            continue


if __name__ == '__main__':
    # вносим в список интересующие вас группы
    # group_list = ['happython', 'python_forum', 'vk_python', 'pirsipy']
    group_list = ['happython']
    parser(group_list)
