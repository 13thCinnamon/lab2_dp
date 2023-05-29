from aiogram import Bot, types
from aiogram.dispatcher import Dispatcher
from aiogram.utils import executor
import xlrd #библиотка

bot = Bot(token='6104916912:AAESYipyosOdHRQfqiGJ0hDV0tICMfCO30A')
dp = Dispatcher(bot)

@dp.message_handler(commands=['start'])
async def process_start_command(message: types.Message):
    await message.reply("Привет, здесь ты можешь получить интресующую тебя информацию по университету\nДля вывода списка команд введи /help")

@dp.message_handler(commands=['help'])
async def process_help_command(message: types.Message):
    await message.reply("/address - узнать адреса и кабинеты интересующих тебя мест\n/map - карта корпусов университета\n/email - узнать эл. адреса преподователей")

@dp.message_handler(commands=['map'])
async def process_help_command(message: types.Message):
    await bot.send_photo(message.from_user.id, types.InputFile('_04042019.jpg'), reply_to_message_id=message.message_id)

@dp.message_handler(commands=['address'])
async def address(message):
    await message.reply('Какое именно место вас интересует?')



@dp.message_handler(commands=['email'])
async def email(message):
    await bot.send_message(message.chat.id, "Введите Фамилию и инициалы преподователя")
    rb = xlrd.open_workbook('bot_db.xls', formatting_info=True)
    sheet = rb.sheet_by_index(0)
    @dp.message_handler(content_types='text')
    async def message_email(message):
        for col_data in sheet.col_values(1):
            if message.text==str(col_data):
                row = str(sheet.cell_value(sheet.col_values(1).index(col_data), 1))
                row += ' - '
                row += str(sheet.cell_value(sheet.col_values(1).index(col_data), 2))
                await bot.send_message(message.chat.id, row)




if __name__ == '__main__':
    executor.start_polling(dp)
