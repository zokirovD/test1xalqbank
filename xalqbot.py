import qrcode
import io
import pandas as pd
import asyncio
import os 
from datetime import datetime, timedelta

from aiogram import Bot, Dispatcher, executor, types
from aiogram.types import ReplyKeyboardMarkup, KeyboardButton, InlineKeyboardButton, InlineKeyboardMarkup
from aiogram.dispatcher.filters import Text
from aiogram.contrib.fsm_storage.memory import MemoryStorage
from aiogram.dispatcher import FSMContext
from aiogram.dispatcher.filters.state import State, StatesGroup
from qrcode.image.pil import PilImage  # PIL versiyasini aniq ko'rsatamiz
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageDraw, ImageFont

# REPLACE WITH YOUR BOT TOKEN
BOT_TOKEN = "7756441821:AAG-I-IjD7H4l_zV6UQL1TfeppyPMZjTK5Y"

# REPLACE WITH YOUR GOOGLE SHEET ID
GOOGLE_SHEET_ID = "1G93-WfwCoI5Sa4J2Dd97Ii8Wi2jovWnNhj4hB3Jpo1k"

# REPLACE WITH PATH TO YOUR GOOGLE CREDENTIALS JSON
GOOGLE_CREDENTIALS_JSON = "testbot-462404-0426d839e8b8.json"

# REPLACE WITH YOUR EXCEL FILE PATH
EMPLOYEE_EXCEL = "employexalqbank.xlsx"

# Load employee data
df_employees = pd.read_excel(EMPLOYEE_EXCEL)
df_employees.columns = df_employees.columns.str.strip().str.lower()
employee_data = {str(row['nps_id']): row for _, row in df_employees.iterrows()}

# Google Sheets setup
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
creds = ServiceAccountCredentials.from_json_keyfile_name(GOOGLE_CREDENTIALS_JSON, scope)
client = gspread.authorize(creds)
sheet = client.open_by_key(GOOGLE_SHEET_ID).sheet1
# Ensure headers exist in the sheet
expected_headers = ['nps_id', 'xodim_fio', "xodim_lavozimi", 'bxm_id', 'telefon_raqam', 'baho', 'sana']
current = sheet.row_values(1)
if current != expected_headers:
    sheet.clear()  # Optional: clear if something else is in the first row
    sheet.append_row(expected_headers)


# Anti-fraud tracking
last_ratings = {}

# Bot setup
bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(bot, storage=MemoryStorage())

class RegisterState(StatesGroup):
    waiting_for_contact = State()
    waiting_for_rating = State()

# Start with argument (used by CUSTOMERS from QR)
@dp.message_handler(lambda message: message.text.startswith("/start ") and message.get_args())
async def handle_qr_start(message: types.Message, state: FSMContext):
    emp_id = message.get_args().strip()
    if emp_id in employee_data:
        await state.update_data(nps_id=emp_id)
        contact_btn = ReplyKeyboardMarkup(resize_keyboard=True).add(
            KeyboardButton("📱 Share Contact", request_contact=True)
        )
        await message.answer("Iltimos, ushbu xodimni baholash uchun kontakt ma'lumotlaringizni ulashing.", reply_markup=contact_btn)
        await RegisterState.waiting_for_contact.set()
    else:
        await message.answer("QR koddagi xodim ID raqami noto'g'ri.")

# Start without arguments (EMPLOYEES use this)
@dp.message_handler(commands=['start'])
async def handle_plain_start(message: types.Message):
    await message.reply("Xush kelibsiz, xodim! Iltimos, QR kodingizni olish uchun NPS raqamingizni yuboring.")

# Employee sends their ID manually

def generate_styled_qr(employee_name: str, url: str) -> io.BytesIO:
    # Generate QR code
    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=10,
        border=2,
    )
    qr.add_data(url)
    qr.make(fit=True)
    qr_img = qr.make_image(fill_color="black", back_color="white").convert('RGB')
    qr_img = qr_img.resize((350, 350))

    # Load logo with verification
    logo = None
    logo_path = "C:/Users/user/Desktop/test1bot/logo.png"  # Make sure this path is correct
    
    try:
        if os.path.exists(logo_path):
            logo = Image.open(logo_path).convert("RGBA")
            print(f"Logo loaded successfully. Original size: {logo.size}")  # Debug print
            logo = logo.resize((250, int(250 * logo.height / logo.width)))
            print(f"Logo resized to: {logo.size}")  # Debug print
        else:
            print(f"Logo file not found at: {os.path.abspath(logo_path)}")  # Debug print
    except Exception as e:
        print(f"Error loading logo: {str(e)}")  # Debug print

    # Load fonts
    try:
        bold_font = ImageFont.truetype("arialbd.ttf", 32)
        regular_font = ImageFont.truetype("arial.ttf",20)
    except:
        print("Using default fonts")  # Debug print
        bold_font = ImageFont.load_default()
        regular_font = bold_font

    # Calculate image dimensions
    width = 500
    height = 40  # Top padding
    
    if logo:
        height += logo.height + 30
    else:
        print("No logo available, adjusting layout")  # Debug print
    
    height += 50 + qr_img.height + 40 + 50 + 40  # Other elements

    # Create image
    img = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(img)
    y_position = 40

    # Add logo if available
    if logo:
        try:
            x_logo = (width - logo.width) // 2
            img.paste(logo, (x_logo, y_position), logo)
            print(f"Logo pasted at position: ({x_logo}, {y_position})")  # Debug print
            y_position += logo.height + 30
        except Exception as e:
            print(f"Error pasting logo: {str(e)}")  # Debug print

    # Add service text
    service_text = "Xizmat sifatini baholang"
    text_width = draw.textlength(service_text, font=bold_font)
    draw.text(((width - text_width) // 2, y_position), 
             service_text, fill="black", font=bold_font)
    y_position += 50

    # Add QR code
    x_qr = (width - qr_img.width) // 2
    img.paste(qr_img, (x_qr, y_position))
    y_position += qr_img.height + 40

    # Add employee name
        # Add employee name (multi-line)
    lines = employee_name.split('\n')
    for line in lines:
        name_width = draw.textlength(line.strip(), font=regular_font)
        draw.text(
            ((width - name_width) // 2, y_position),
            line.strip(),
            fill="black",
            font=regular_font
        )
        y_position += 35  # Space between lines



    # Final output
    buf = io.BytesIO()
    img.save(buf, format="PNG", quality=100)
    buf.seek(0)
    return buf



# ==== QR Code Handler ====
@dp.message_handler(lambda message: message.text.isdigit())
async def register_employee(message: types.Message):
    emp_id = message.text.strip()
    if emp_id in employee_data:
        emp_info = employee_data[emp_id]
        qr_url = f"https://t.me/{(await bot.get_me()).username}?start={emp_id}"
        qr_buf = generate_styled_qr(emp_info['xodim_fio'], qr_url)
        await bot.send_photo(message.chat.id, qr_buf, caption="QR kodni chiqarib olasiz!")
    else:
        await message.reply("NPS raqam topilmadi!")


# Customer shares contact
@dp.message_handler(content_types=types.ContentType.CONTACT, state=RegisterState.waiting_for_contact)
async def process_contact(message: types.Message, state: FSMContext):
    phone = message.contact.phone_number
    data = await state.get_data()
    emp_id = data.get("nps_id")

    # Anti-fraud check
    key = (phone, emp_id)
    now = datetime.now()
    if key in last_ratings and now - last_ratings[key] < timedelta(hours=1):
        await message.answer("⛔ Bitta xodimni 1 soat ichida 1 marta baholay olasiz.")
        await state.finish()
        return
    last_ratings[key] = now

    emp_info = employee_data[emp_id]
    await state.update_data(phone=phone)

    buttons = [
        InlineKeyboardButton("👍 Zo'r", callback_data="Zo'r"),
        InlineKeyboardButton("😐 Yaxshi", callback_data="Yaxshi"),
        InlineKeyboardButton("👎 Yomon", callback_data="Yomon")
    ]
    markup = InlineKeyboardMarkup().add(*buttons)

    msg = f"Siz baholayapsiz:\n<b>{emp_info['xodim_fio']}</b>\nxodim_lavozimi: {emp_info['xodim_lavozimi']}"
    await message.answer(msg, parse_mode="HTML")
    await message.answer("Bahoni tanlang:", reply_markup=markup)
    await RegisterState.waiting_for_rating.set()

# Customer selects rating
@dp.callback_query_handler(lambda c: c.data in ["Zo'r", "Yaxshi", "Yomon"], state=RegisterState.waiting_for_rating)
async def process_rating(callback: types.CallbackQuery, state: FSMContext):
    data = await state.get_data()
    emp_id = data["nps_id"]
    phone = data["phone"]
    rating = callback.data
    emp_info = employee_data[emp_id]

    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    # Google Sheets write
    sheet.append_row([
        emp_id,
        emp_info['xodim_fio'],
        emp_info["xodim_lavozimi"],
        emp_info.get('bxm_id', ''),
        phone,
        rating,
        now_str
    ])

    # Delete inline buttons
    await callback.message.edit_reply_markup(reply_markup=None)
     # Send thank-you message
    await callback.message.answer("✅ Bizni baholaganingiz uchun rahmat!")
    await state.finish()

# Run bot
if __name__ == '__main__':
    executor.start_polling(dp, skip_updates=True)