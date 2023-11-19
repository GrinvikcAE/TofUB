import asyncio
import os
from datetime import datetime
from time import sleep
from threading import Thread

from aiogram import Bot, Dispatcher
from aiogram.filters import Command, CommandObject
from aiogram.types import Message
from aiogram.fsm.storage.memory import MemoryStorage
from bestconfig import Config

from TofUB import (create_service_account, get_list_of_files, check_file, check_tasks, sstart,
                   add_permission, delete_permission, delete_file)

config = Config('./keys/settings.json')
TOKEN_BOT = config.get('TOKEN_BOT')
EMAILS = config.get('EMAILS')
USERNAMES = config.get('USERNAMES')
http_auth = create_service_account()
bot = Bot(TOKEN_BOT, parse_mode='HTML')
dp = Dispatcher(storage=MemoryStorage())


def day():
    names_of_files = []
    emails = []
    with open('./data/start.txt', 'r', encoding="UTF-8") as f:
        for line in f:
            names_of_files.append(line.strip('\n').split('\t')[2])
            emails.append(line.strip('\n').split('\t')[1])

    # sstart(http_auth, names_of_files, emails)
    # sleep(600)
    q = names_of_files[0][:5]
    stop_len = len(get_list_of_files(http_auth, q=q)['files'])
    i = 0
    name_files = []

    while len(name_files) != stop_len:
        i += 1
        print('Check ### ' + str(i) + ' ' + str(datetime.now()))
        temp = check_file(http_auth, q, name_files)
        print(temp)
        for j in range(len(temp)):
            if temp[j] not in name_files:
                name_files.append(temp[j])
        sleep(60)
    print('Check completed')


@dp.message(Command('start'))
async def start(message: Message):
    await message.answer(f'Hello, <b>{message.from_user.first_name}</b>')


@dp.message(Command('help'))
async def call_help(message: Message):
    await message.answer(f'/get - Список файлов и разрешения для них\n'
                         f'/add - Добавить разрешения: email file_id\n'
                         f'/del - Удалить разрешения: email file_name\n'
                         f'/create - Добавить в .txt файл для последующей автоматической выдачи разрешения: '
                         f'ФамилияИО email file_name\n'
                         f'/refresh - удалить start.txt\n'
                         f'/start_day - Начать выдачу разрешений и проверку файлов на завершение\n'
                         f'/del_file - Удалить файл: file_id')


@dp.message(Command('get'))
async def get_list(message: Message, command: CommandObject):
    if message.from_user.username in USERNAMES:
        if command.args is None:
            q = None
        else:
            q = command.args
        temp = get_list_of_files(http_auth, q=q)

        result = {}
        for i in range(len(temp['files'])):
            result[temp['files'][i]['name']] = [temp['files'][i]['mimeType'], temp['files'][i]['id']]
            for j in range(len(temp['files'][i]['permissions'])):
                try:
                    if q is not None and temp['files'][i]['permissions'][j]['emailAddress'] not in EMAILS:
                        result[temp['files'][i]['name']].append(temp['files'][i]['permissions'][j]['emailAddress'])
                except Exception as e:
                    print(f'Not found {e}')

        for file in result:
            await message.answer(f'Filename: <b>{file}</b>, \n{result[file]}')
    else:
        await message.reply(f"You don't have permission to this command")


@dp.message(Command('add'))
async def add(message: Message, command: CommandObject):
    if message.from_user.username in USERNAMES:
        email, file_id = command.args.split(' ')
        await message.reply(f'Give permission to {email}, file {file_id}')
        add_permission(http_auth, email, file_id)
    else:
        await message.reply(f"You don't have permission to this command")


@dp.message(Command('del'))
async def delete(message: Message, command: CommandObject):
    if message.from_user.username in USERNAMES:
        email, q = command.args.split(' ')
        delete_permission(http_auth, email=email, q=q)
        await message.reply(f'Delete permission from {email}, file {q}')
    else:
        await message.reply(f"You don't have permission to this command")


@dp.message(Command('create'))
async def create(message: Message, command: CommandObject):
    if message.from_user.username in USERNAMES:
        with open('./data/start.txt', 'a', encoding="UTF-8") as f:
            lst = command.args.split(' ')
            f.write(lst[0] + '\t' + lst[1] + '\t' + lst[2] + '\n')
    else:
        await message.reply(f"You don't have permission to this command")


@dp.message(Command('refresh'))
async def refresh(message: Message):
    if message.from_user.username in USERNAMES:
        os.remove('./data/start.txt')
        file = open('./data/start.txt', 'w', encoding="UTF-8")
        file.close()
    else:
        await message.reply(f"You don't have permission to this command")


@dp.message(Command('start_day'))
async def start_day(message: Message):
    if message.from_user.username in USERNAMES:
        th = Thread(target=day).start()
    else:
        await message.reply(f"You don't have permission to this command")


@dp.message(Command('check'))
async def check(message: Message, command: CommandObject):
    if message.from_user.username in USERNAMES:
        if command.args is None:
            q = None
        else:
            q = command.args
        check_tasks(http_auth, q=q)
    else:
        await message.reply(f"You don't have permission to this command")


@dp.message(Command('del_file'))
async def del_file(message: Message, command: CommandObject):
    if message.from_user.username in USERNAMES:
        delete_file(http_auth, file_id=command.args)
    else:
        await message.reply(f"You don't have permission to this command")


'''
@dp.message()
async def echo(message: Message):
    await message.reply(f'No command')
'''


async def main():
    await bot.delete_webhook(drop_pending_updates=True)
    await dp.start_polling(bot, skip_updates=True)


if __name__ == '__main__':
    asyncio.run(main())
