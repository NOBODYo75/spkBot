import asyncio

import pandas as pd
from aiogram import Bot, Dispatcher, types
from aiogram.types import ParseMode

API_TOKEN = '7474508608:AAFlexXt3-wKR-9qdZV6y8IRHLe76dcH3Jo'  # Замените на ваш API токен

# Инициализация бота и диспетчера
bot = Bot(token=API_TOKEN)
dp = Dispatcher(bot)


@dp.message_handler(commands=['start'])
async def start(message):
    await message.answer(
        "Привет! Я бот для удобного просматривания расписания для Северского Промышленного колледжа. Для того чтобы узнать расписание напиши неделюю, группу, день. Пример 'Нечётная Д192 Понедельник' ")


async def odd_D192_monday(message: types.Message):
    # Загрузка файла Excel и извлечение нужных данных
    rasp = pd.read_excel(
        io='rasp.xlsx',  # Укажите путь к вашему файлу
        engine='openpyxl',
        usecols='EB:EE',  # Укажите нужные столбцы
        header=6,  # Укажите нужную строку заголовка
        nrows=11  # Укажите количество строк для чтения
    )

    # Заполнение пустых значений и приведение всех значений к строковому типу
    rasp = rasp.fillna('').astype(str)

    # Переименование столбцов и индексов
    rasp.columns = ['Предмет1', 'Кабинет1', 'Предмет2', 'Кабинет2']
    rasp.index = ['1', '', '2', '', '3', '', '4', '', '5', '', '6']

    # Формируем итоговое сообщение
    final_message_text = ""

    # Заполнение сообщения для каждой пары строк
    for i in range(0, len(rasp), 2):
        # Извлекаем данные для текущей и следующей строки (если есть)
        row1 = rasp.iloc[i]
        row2 = rasp.iloc[i + 1] if i + 1 < len(rasp) else pd.Series(['', '', '', ''],
                                                                    index=['Предмет1', 'Кабинет1', 'Предмет2',
                                                                           'Кабинет2'])

        # Проверяем, что хотя бы одно значение не пустое для текущей пары строк
        if any(row1[['Предмет1', 'Предмет2']].values != '') or any(row2[['Предмет1', 'Предмет2']].values != ''):
            # Формируем текст сообщения для текущей пары строк
            message_text = f"#{rasp.index[i]}:\n"

            # Добавляем информацию о предмете и кабинете из первого столбца
            if row1['Предмет1'] != '':
                if row1['Кабинет1'] != '':
                    message_text += "<i>Пара у первой подгруппы</i>\n"
                    message_text += f"<b>Предмет:</b> {row1['Предмет1']}\n"
                    message_text += f"<b>Кабинет:</b> {row1['Кабинет1']}\n"
                    message_text += f"<b>Преподаватель:</b> {row2['Предмет1']}\n"
                    message_text += "\n"
                elif row1['Кабинет2'] != '':
                    message_text += "<i>У всей группы</i>\n"
                    message_text += f"<b>Предмет:</b> {row1['Предмет1']}\n"
                    message_text += f"<b>Кабинет:</b> {row1['Кабинет2']}\n"
                    message_text += f"<b>Преподаватель:</b> {row2['Предмет1']}\n"
                    message_text += "\n"
                else:
                    message_text += f"<b>Предмет:</b> {row1['Предмет1']}\n"
                    message_text += "\n"

            # Добавляем информацию о предмете и кабинете из второго столбца
            if row1['Предмет2'] != '':
                message_text += "<i>Пара у второй подгруппы</i>\n"
                message_text += f"<b>Предмет:</b> {row1['Предмет2']}\n"
                if row1['Кабинет2'] != '':
                    message_text += f"<b>Кабинет:</b> {row1['Кабинет2']}\n"
                    message_text += f"<b>Преподаватель:</b> {row2['Предмет2']}\n"
                message_text += "\n"

                # Добавляем информацию о преподавателе и кабинете из первого столбца

                if row2['Кабинет1'] != '':
                    message_text += f"<b>Кабинет:</b> {row2['Кабинет1']}\n"
                message_text += "\n"

            # Добавляем информацию о преподавателе и кабинете из второго столбца
            if row2['Предмет2'] != '':
                if row2['Кабинет2'] != '':
                    message_text += f"<b>Кабинет:</b> {row2['Кабинет2']}\n"
                message_text += "\n"

            # Добавляем сформированный текст к итоговому сообщению
            final_message_text += message_text + "\n"

    # Отправка итогового сообщения пользователю с HTML форматированием
    await message.answer(final_message_text, parse_mode=ParseMode.HTML)


async def odd_D192_tuesday(message: types.Message):
    # Загрузка файла Excel и извлечение нужных данных
    rasp = pd.read_excel(
        io='rasp.xlsx',  # Укажите путь к вашему файлу
        engine='openpyxl',
        usecols='EB:EE',  # Укажите нужные столбцы
        header=18,  # Укажите нужную строку заголовка
        nrows=11  # Укажите количество строк для чтения
    )

    # Заполнение пустых значений и приведение всех значений к строковому типу
    rasp = rasp.fillna('').astype(str)

    # Переименование столбцов и индексов
    rasp.columns = ['Предмет1', 'Кабинет1', 'Предмет2', 'Кабинет2']
    rasp.index = ['1', '', '2', '', '3', '', '4', '', '5', '', '6']

    # Формируем итоговое сообщение
    final_message_text = ""

    # Заполнение сообщения для каждой пары строк
    for i in range(0, len(rasp), 2):
        # Извлекаем данные для текущей и следующей строки (если есть)
        row1 = rasp.iloc[i]
        row2 = rasp.iloc[i + 1] if i + 1 < len(rasp) else pd.Series(['', '', '', ''],
                                                                    index=['Предмет1', 'Кабинет1', 'Предмет2',
                                                                           'Кабинет2'])

        # Проверяем, что хотя бы одно значение не пустое для текущей пары строк
        if any(row1[['Предмет1', 'Предмет2']].values != '') or any(row2[['Предмет1', 'Предмет2']].values != ''):
            # Формируем текст сообщения для текущей пары строк
            message_text = f"#{rasp.index[i]}:\n"

            # Добавляем информацию о предмете и кабинете из первого столбца
            if row1['Предмет1'] != '':
                if row1['Кабинет1'] != '':
                    message_text += "<i>Пара у первой подгруппы</i>\n"
                    message_text += f"<b>Предмет:</b> {row1['Предмет1']}\n"
                    message_text += f"<b>Кабинет:</b> {row1['Кабинет1']}\n"
                    message_text += f"<b>Преподаватель:</b> {row2['Предмет1']}\n"
                    message_text += "\n"
                elif row1['Кабинет2'] != '':
                    message_text += "<i>У всей группы</i>\n"
                    message_text += f"<b>Предмет:</b> {row1['Предмет1']}\n"
                    message_text += f"<b>Кабинет:</b> {row1['Кабинет2']}\n"
                    message_text += f"<b>Преподаватель:</b> {row2['Предмет1']}\n"
                    message_text += "\n"
                else:
                    message_text += f"<b>Предмет:</b> {row1['Предмет1']}\n"
                    message_text += "\n"

            # Добавляем информацию о предмете и кабинете из второго столбца
            if row1['Предмет2'] != '':
                message_text += "<i>Пара у второй подгруппы</i>\n"
                message_text += f"<b>Предмет:</b> {row1['Предмет2']}\n"
                if row1['Кабинет2'] != '':
                    message_text += f"<b>Кабинет:</b> {row1['Кабинет2']}\n"
                    message_text += f"<b>Преподаватель:</b> {row2['Предмет2']}\n"
                message_text += "\n"

                # Добавляем информацию о преподавателе и кабинете из первого столбца

                if row2['Кабинет1'] != '':
                    message_text += f"<b>Кабинет:</b> {row2['Кабинет1']}\n"
                message_text += "\n"

            # Добавляем информацию о преподавателе и кабинете из второго столбца
            if row2['Предмет2'] != '':
                if row2['Кабинет2'] != '':
                    message_text += f"<b>Кабинет:</b> {row2['Кабинет2']}\n"
                message_text += "\n"

            # Добавляем сформированный текст к итоговому сообщению
            final_message_text += message_text + "\n"

    # Отправка итогового сообщения пользователю с HTML форматированием
    await message.answer(final_message_text, parse_mode=ParseMode.HTML)


async def odd_D192_wednesday(message: types.Message):
    # Загрузка файла Excel и извлечение нужных данных
    rasp = pd.read_excel(
        io='rasp.xlsx',  # Укажите путь к вашему файлу
        engine='openpyxl',
        usecols='EB:EE',  # Укажите нужные столбцы
        header=30,  # Укажите нужную строку заголовка
        nrows=11  # Укажите количество строк для чтения
    )

    # Заполнение пустых значений и приведение всех значений к строковому типу
    rasp = rasp.fillna('').astype(str)

    # Переименование столбцов и индексов
    rasp.columns = ['Предмет1', 'Кабинет1', 'Предмет2', 'Кабинет2']
    rasp.index = ['1', '', '2', '', '3', '', '4', '', '5', '', '6']

    # Формируем итоговое сообщение
    final_message_text = ""

    # Заполнение сообщения для каждой пары строк
    for i in range(0, len(rasp), 2):
        # Извлекаем данные для текущей и следующей строки (если есть)
        row1 = rasp.iloc[i]
        row2 = rasp.iloc[i + 1] if i + 1 < len(rasp) else pd.Series(['', '', '', ''],
                                                                    index=['Предмет1', 'Кабинет1', 'Предмет2',
                                                                           'Кабинет2'])

        # Проверяем, что хотя бы одно значение не пустое для текущей пары строк
        if any(row1[['Предмет1', 'Предмет2']].values != '') or any(row2[['Предмет1', 'Предмет2']].values != ''):
            # Формируем текст сообщения для текущей пары строк
            message_text = f"#{rasp.index[i]}:\n"

            # Добавляем информацию о предмете и кабинете из первого столбца
            if row1['Предмет1'] != '':
                if row1['Кабинет1'] != '':
                    message_text += "<i>Пара у первой подгруппы</i>\n"
                    message_text += f"<b>Предмет:</b> {row1['Предмет1']}\n"
                    message_text += f"<b>Кабинет:</b> {row1['Кабинет1']}\n"
                    message_text += f"<b>Преподаватель:</b> {row2['Предмет1']}\n"
                    message_text += "\n"
                elif row1['Кабинет2'] != '':
                    message_text += "<i>У всей группы</i>\n"
                    message_text += f"<b>Предмет:</b> {row1['Предмет1']}\n"
                    message_text += f"<b>Кабинет:</b> {row1['Кабинет2']}\n"
                    message_text += f"<b>Преподаватель:</b> {row2['Предмет1']}\n"
                    message_text += "\n"
                else:
                    message_text += f"<b>Предмет:</b> {row1['Предмет1']}\n"
                    message_text += "\n"

            # Добавляем информацию о предмете и кабинете из второго столбца
            if row1['Предмет2'] != '':
                message_text += "<i>Пара у второй подгруппы</i>\n"
                message_text += f"<b>Предмет:</b> {row1['Предмет2']}\n"
                if row1['Кабинет2'] != '':
                    message_text += f"<b>Кабинет:</b> {row1['Кабинет2']}\n"
                    message_text += f"<b>Преподаватель:</b> {row2['Предмет2']}\n"
                message_text += "\n"

                # Добавляем информацию о преподавателе и кабинете из первого столбца

                if row2['Кабинет1'] != '':
                    message_text += f"<b>Кабинет:</b> {row2['Кабинет1']}\n"
                message_text += "\n"

            # Добавляем информацию о преподавателе и кабинете из второго столбца
            if row2['Предмет2'] != '':
                if row2['Кабинет2'] != '':
                    message_text += f"<b>Кабинет:</b> {row2['Кабинет2']}\n"
                message_text += "\n"

            # Добавляем сформированный текст к итоговому сообщению
            final_message_text += message_text + "\n"

    # Отправка итогового сообщения пользователю с HTML форматированием
    await message.answer(final_message_text, parse_mode=ParseMode.HTML)


async def main():
    # Регистрируем обработчики для каждого дня недели
    dp.register_message_handler(odd_D192_monday, text=["Нечётная Д192 понедельник"])
    dp.register_message_handler(odd_D192_tuesday, text=["Нечётная Д192 вторник"])
    dp.register_message_handler(odd_D192_wednesday, text=["Нечётная Д192 среда"])

    await dp.start_polling()

if __name__ == "__main__":
    asyncio.run(main())
