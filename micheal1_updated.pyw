import csv
from datetime import datetime, timedelta, timezone
import io
import os
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk
import pandas as pd
import requests
from selenium import webdriver
from selenium.webdriver.chrome.options import Options

# استدعاء مكتبة الصورة لأخذ الاسكرين شوت
try:
  from PIL import ImageGrab
except ImportError:
  ImageGrab = None

# ====================================================================
# ⚙️ الإعدادات العامة وقاموس المشرفين
# ====================================================================
EGYPT_TZ = timezone(timedelta(hours=3))

if getattr(sys, 'frozen', False):
  SCRIPT_DIR = os.path.dirname(sys.executable)
else:
  SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))

# قاموس عام لتخزين بيانات المشرفين { 'Rider_ID': 'اسم المشرف' }
supervisors_map = {}


def upload_supervisors_file(parent_win=None):
  """دالة لرفع شيت الإكسيل وقراءة ID المندوب (عمود A) واسم المشرف (عمود B)"""
  global supervisors_map
  filepath = filedialog.askopenfilename(
      title='اختر ملف إكسيل أو CSV للمشرفين',
      filetypes=[
          ('Excel / CSV Files', '*.xlsx *.xls *.csv'),
          ('All Files', '*.*'),
      ],
  )
  if not filepath:
    return

  try:
    if filepath.endswith('.csv'):
      df = pd.read_csv(filepath, header=None)
    else:
      df = pd.read_excel(filepath, header=None)

    raw_map = dict(
        zip(df[0].astype(str).str.strip(), df[1].astype(str).str.strip())
    )

    supervisors_map = {
        k: v
        for k, v in raw_map.items()
        if k.lower()
        not in ['id', 'rider id', 'الاي دي', 'المندوب', 'nan', 'none', '']
    }

    messagebox.showinfo(
        'تم التحميل بنجاح',
        f'تم ربط بيانات {len(supervisors_map)} مندوب بالمشرفين بنجاح!',
    )

    if parent_win:
      view_supervisors_popup(parent_win)

  except Exception as e:
    messagebox.showerror(
        'خطأ في التحميل',
        f'فشل قراءة الملف: {e}\nتأكد من تثبيت pandas و openpyxl',
    )


def view_supervisors_popup(parent):
  """نافذة لمعاينة المشرفين المرفوعين مع شريط بحث وإمكانية الحذف"""
  if not supervisors_map:
    messagebox.showinfo(
        'تنبيه', 'لم يتم رفع أي شيت مشرفين حتى الآن، القائمة فارغة!'
    )
    return

  popup = tk.Toplevel(parent)
  popup.title(f'قائمة المشرفين المرفوعة ({len(supervisors_map)} مندوب)')
  popup.geometry('550x580')

  top_bar = tk.Frame(popup, padx=10, pady=8)
  top_bar.pack(fill='x')

  lbl_count = tk.Label(
      top_bar,
      text=f'📊 إجمالي المناديب المربوطين: {len(supervisors_map)}',
      font=('Arial', 10, 'bold'),
      fg='#059669',
  )
  lbl_count.pack(side='top', anchor='w', pady=(0, 5))

  search_frame = tk.Frame(top_bar)
  search_frame.pack(fill='x')

  tk.Label(search_frame, text='🔍 بحث (ID / مشرف):', font=('Arial', 10)).pack(
      side='left', padx=2
  )
  ent_search = tk.Entry(search_frame, font=('Arial', 10))
  ent_search.pack(side='left', fill='x', expand=True, padx=5)

  cols = ('ID المندوب', 'اسم المشرف')
  tree = ttk.Treeview(popup, columns=cols, show='headings', height=16)
  tree.heading('ID المندوب', text='ID المندوب')
  tree.heading('اسم المشرف', text='اسم المشرف')

  tree.column('ID المندوب', width=160, anchor='center')
  tree.column('اسم المشرف', width=280, anchor='center')

  vsb = ttk.Scrollbar(popup, orient='vertical', command=tree.yview)
  tree.configure(yscrollcommand=vsb.set)

  vsb.pack(side='right', fill='y')
  tree.pack(fill='both', expand=True, padx=10, pady=5)

  def refresh_tree(*args):
    for item in tree.get_children():
      tree.delete(item)
    query = ent_search.get().strip().lower()

    for rider_id, sup_name in supervisors_map.items():
      if query and not (
          query in str(rider_id).lower() or query in str(sup_name).lower()
      ):
        continue
      tree.insert('', 'end', values=(rider_id, sup_name))

  ent_search.bind('<KeyRelease>', refresh_tree)
  refresh_tree()

  btn_bar = tk.Frame(popup, padx=10, pady=10)
  btn_bar.pack(fill='x')

  def clear_data():
    if messagebox.askyesno(
        'تأكيد المسح', 'هل أنت تأكد من مسح جميع بيانات المشرفين المرفوعة؟'
    ):
      supervisors_map.clear()
      messagebox.showinfo('تم المسح', 'تم تفريغ قائمة المشرفين بنجاح.')
      popup.destroy()

  btn_clear = tk.Button(
      btn_bar,
      text='🗑️ مسح القائمة الحالية',
      bg='#dc2626',
      fg='white',
      font=('Arial', 9, 'bold'),
      command=clear_data,
  )
  btn_clear.pack(side='left')

  btn_close = tk.Button(
      btn_bar,
      text='إغلاق',
      font=('Arial', 9),
      width=10,
      command=popup.destroy,
  )
  btn_close.pack(side='right')


# ====================================================================
# 🛠️ أدوات التصدير، النسخ، والتصوير
# ====================================================================
def export_treeview_to_csv(tree, default_filename='Riders_Export.csv'):
  filepath = filedialog.asksaveasfilename(
      defaultextension='.csv',
      filetypes=[('CSV Excel File', '*.csv'), ('All Files', '*.*')],
      initialfile=default_filename,
  )
  if not filepath:
    return
  try:
    columns = list(tree['columns'])
    with open(filepath, 'w', newline='', encoding='utf-8-sig') as f:
      writer = csv.writer(f)
      writer.writerow(columns)
      for row_id in tree.get_children():
        writer.writerow(tree.item(row_id)['values'])
    messagebox.showinfo('نجاح التصدير', 'تم تصدير البيانات إلى ملف إكسيل بنجاح!')
  except Exception as e:
    messagebox.showerror('خطأ', f'حدث خطأ أثناء التصدير: {e}')


def copy_treeview_to_clipboard(root_win, tree):
  """نسخ البيانات الظاهرة بالجدول إلى الحافظة"""
  try:
    columns = list(tree['columns'])
    rows = [columns]
    for row_id in tree.get_children():
      vals = tree.item(row_id)['values']
      if vals:
        rows.append([str(v) for v in vals])

    text_data = '\n'.join(['\t'.join(row) for row in rows])
    root_win.clipboard_clear()
    root_win.clipboard_append(text_data)
    root_win.update()
    messagebox.showinfo('تم النسخ', 'تم نسخ بيانات الجدول بنجاح!')
  except Exception as e:
    messagebox.showerror('خطأ', f'فشل نسخ البيانات: {e}')


def take_window_screenshot(widget, title='Screenshot'):
  if ImageGrab is None:
    messagebox.showerror(
        'المكتبة غير متوفرة',
        'يرجى تثبيت مكتبة Pillow لأخذ اسكرين شوت عبر أمر:\npip install Pillow',
    )
    return
  try:
    widget.update_idletasks()
    x = widget.winfo_rootx()
    y = widget.winfo_rooty()
    w = widget.winfo_width()
    h = widget.winfo_height()

    img = ImageGrab.grab(bbox=(x, y, x + w, y + h))

    filename = f"{title}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    filepath = filedialog.asksaveasfilename(
        defaultextension='.png',
        filetypes=[('PNG Image', '*.png'), ('All Files', '*.*')],
        initialfile=filename,
    )
    if filepath:
      img.save(filepath)
      messagebox.showinfo('تم الحفظ', 'تم أخذ لقطة الشاشة للجدول وحفظها بنجاح!')
  except Exception as e:
    messagebox.showerror('خطأ', f'فشل أخذ لقطة الشاشة: {e}')


# ====================================================================
# 🌐 مدير المتصفح الموحد (Shared Browser Manager)
# ====================================================================
class SharedBrowser:

  def __init__(self):
    self.driver = None
    self.lock = threading.RLock()
    self.is_logged_in_printed = False

  def get_driver(self, logger_func=None):
    with self.lock:
      if self.driver is None:
        try:
          profile_dir = os.path.join(SCRIPT_DIR, 'profile_shared_master')

          chrome_options = Options()
          chrome_options.add_argument(f'--user-data-dir={profile_dir}')
          chrome_options.add_argument('--no-sandbox')
          chrome_options.add_argument('--disable-dev-shm-usage')
          chrome_options.add_argument(
              '--disable-blink-features=AutomationControlled'
          )
          chrome_options.add_experimental_option(
              'excludeSwitches', ['enable-automation']
          )
          chrome_options.add_experimental_option(
              'useAutomationExtension', False
          )
          if logger_func:
            logger_func('⚙️ جاري فتح نافذة المتصفح الموحد...')

          self.driver = webdriver.Chrome(options=chrome_options)
          self.driver.execute_cdp_cmd(
              'Page.addScriptToEvaluateOnNewDocument', {
                  'source': (
                      'Object.defineProperty(navigator, "webdriver", {get: ()'
                      ' => undefined})'
                  )
              },
          )
          self.driver.get(
              'https://eg.me.logisticsbackoffice.com/dashboard/rooster/live-3pl'
          )
        except Exception as e:
          if logger_func:
            logger_func(f'❌ فشل فتح المتصفح! : {e}')
          self.driver = None
      return self.driver

  def get_auth_and_cookies(self, logger_func=None):
    with self.lock:
      drv = self.get_driver(logger_func)
      if not drv:
        return None, None

      try:
        cookies_list = drv.get_cookies()
        dhh_token = next(
            (
                c['value']
                for c in cookies_list
                if c['name'] in ['dhh_token', 'access_token', 'token']
            ),
            None,
        )

        if dhh_token:
          if not self.is_logged_in_printed:
            if logger_func:
              logger_func('⚡ الجلسة نشطة! تم سحب التوكن وإخفاء المتصفح...')
            self.is_logged_in_printed = True
            try:
              drv.minimize_window()
            except Exception:
              pass

          cookie_str = '; '.join(
              [f"{c['name']}={c['value']}" for c in cookies_list]
          )
          return f'Bearer {dhh_token}', cookie_str
        else:
          if logger_func and not self.is_logged_in_printed:
            logger_func('🔒 يرجى تسجيل الدخول داخل نافذة الكروم المفتوحة...')
            self.is_logged_in_printed = True

      except Exception as ce:
        err_msg = str(ce).lower()
        if 'invalid session id' in err_msg or 'no such window' in err_msg:
          if logger_func:
            logger_func('⚠️ تم إغلاق المتصفح الموحد، جاري إعادة الفتح...')
          try:
            drv.quit()
          except Exception:
            pass
          self.driver = None
          self.is_logged_in_printed = False

      return None, None

  def refresh_page(self):
    with self.lock:
      if self.driver:
        try:
          self.driver.refresh()
          self.is_logged_in_printed = False
        except Exception:
          pass

  def close(self):
    with self.lock:
      if self.driver:
        try:
          self.driver.quit()
        except Exception:
          pass
        self.driver = None


# ====================================================================
# 🛠️ الدوال المساعدة المشتركة
# ====================================================================
def parse_api_datetime(dt_str):
  if not dt_str:
    return None
  try:
    clean_str = str(dt_str).replace('Z', '+00:00').replace(' ', 'T')
    dt = datetime.fromisoformat(clean_str)
    if dt.tzinfo is None:
      dt = dt.replace(tzinfo=EGYPT_TZ)
    return dt.astimezone(EGYPT_TZ)
  except Exception:
    return None


def seconds_to_hms_live(d):
  try:
    d = int(float(d))
    if d < 0:
      return '0:00:00'
    h = d // 3600
    m = (d % 3600) // 60
    s = d % 3600 % 60
    return f'{h}:{m:02d}:{s:02d}'
  except Exception:
    return '0:00:00'


def format_to_time_only_live(date_string):
  if not date_string:
    return ''
  try:
    dt_obj = parse_api_datetime(date_string)
    if dt_obj:
      return dt_obj.strftime('%H:%M')
    return date_string
  except Exception:
    return date_string


def calculate_duration_hms(start_str, end_str):
  if not start_str or not end_str:
    return '0:00'
  try:
    d1 = datetime.fromisoformat(str(start_str).replace('Z', '+00:00'))
    d2 = datetime.fromisoformat(str(end_str).replace('Z', '+00:00'))
    diff = d2 - d1
    diff_mins = int(diff.total_seconds() / 60)
    if diff_mins < 0:
      return '0:00'
    hrs = diff_mins // 60
    mins = diff_mins % 60
    return f'{hrs}:{mins:02d}'
  except Exception:
    return '0:00'


# ====================================================================
# 🔄 كلاس المحرك اللوجستي
# ====================================================================
class ScraperEngine:

  def __init__(
      self,
      name,
      target_cities,
      sheet_url,
      csv_filename,
      shared_browser,
      live_interval=120,
      booking_interval=300,
      days_ahead=7,
      log_widget=None,
      status_label=None,
  ):
    self.name = name
    self.target_cities = target_cities
    self.sheet_url = sheet_url
    self.csv_filename = csv_filename
    self.shared_browser = shared_browser
    self.log_widget = log_widget
    self.status_label = status_label

    self.is_running = False
    self.thread = None

    self.mode = 'BOTH'
    self.manual_live_trigger = False
    self.manual_booking_trigger = False

    self.live_interval = live_interval
    self.booking_interval = booking_interval
    self.days_ahead = days_ahead

    self.latest_live_data = []
    self.latest_booking_data = []
    self.on_live_updated = None

  def log(self, text):
    timestamp = datetime.now(EGYPT_TZ).strftime('%H:%M:%S')
    msg = f'[{timestamp}] {text}\n'

    def _update():
      if self.log_widget:
        self.log_widget.configure(state='normal')
        self.log_widget.insert(tk.END, msg)
        self.log_widget.see(tk.END)
        self.log_widget.configure(state='disabled')

    if self.log_widget:
      self.log_widget.after(0, _update)

  def update_status(self, text, color='black'):
    def _update():
      if self.status_label:
        self.status_label.config(text=f'الحالة: {text}', fg=color)

    if self.status_label:
      self.status_label.after(0, _update)

  def trigger_manual_live(self):
    if not self.is_running:
      messagebox.showwarning(
          'تنبيه', 'برجاء تشغيل القسم أولاً بالضغط على زر ▶️ تشغيل.'
      )
      return
    self.manual_live_trigger = True
    self.log('⚡ تم إرسال أمر لتحديث اللايف فوراً بتكة زرار...')

  def trigger_manual_booking(self):
    if not self.is_running:
      messagebox.showwarning(
          'تنبيه', 'برجاء تشغيل القسم أولاً بالضغط على زر ▶️ تشغيل.'
      )
      return
    self.manual_booking_trigger = True
    self.log('⚡ تم إرسال أمر لتحديث الحجز فوراً بتكة زرار...')

  def send_data_to_sheet(self, data_type, data, max_retries=3):
    for attempt in range(1, max_retries + 1):
      try:
        response = requests.post(
            self.sheet_url,
            json={'type': data_type, 'data': data},
            timeout=120,
        )
        if response.status_code == 200:
          return True
      except requests.exceptions.Timeout:
        self.log(
            f'⚠️ تأخر رد جوجل لشيت ({data_type.upper()})، جاري إعادة'
            f' المحاولة ({attempt}/{max_retries})...'
        )
      except Exception as e:
        self.log(
            f'⚠️ خطأ أثناء إرسال البيانات ({data_type.upper()}):'
            f' {e} (محاولة {attempt}/{max_retries})'
        )
      time.sleep(3)
    return False

  def start(self):
    if not self.is_running:
      self.is_running = True
      self.update_status('جاري التشغيل...', 'green')
      self.thread = threading.Thread(
          target=self.run_loop, name=f'Thread-{self.name}', daemon=True
      )
      self.thread.start()

  def stop(self):
    if self.is_running:
      self.is_running = False
      self.update_status('جاري الإيقاف...', 'orange')
      self.log('🛑 تم إيقاف عملية هذا القسم.')

  def run_loop(self):
    last_live_run = 0
    last_booking_run = 0

    while self.is_running:
      try:
        current_timestamp = time.time()
        time_to_live = max(
            0, int(self.live_interval - (current_timestamp - last_live_run))
        )

        # 1️⃣ LIVE CYCLE
        is_scheduled_live = (self.mode in ['LIVE', 'BOTH']) and (
            time_to_live <= 0
        )
        is_manual_live = self.manual_live_trigger

        if (is_scheduled_live or is_manual_live) and self.is_running:
          self.manual_live_trigger = False

          auth_token, cookie = self.shared_browser.get_auth_and_cookies(
              self.log
          )
          if not self.is_running or not auth_token:
            time.sleep(3)
            continue

          headers = {
              'Authorization': auth_token,
              'Cookie': cookie,
              'Accept': 'application/json',
              'User-Agent': 'Mozilla/5.0',
          }
          current_time_live = datetime.now(EGYPT_TZ).strftime(
              '%Y-%m-%d %I:%M:%S %p'
          )
          self.log(f'🚀 بدء دورة سحب LIVE (كل {self.live_interval} ثانية)...')

          all_riders = []
          session_expired = False

          for city in self.target_cities:
            if not self.is_running:
              break
            page = 0
            has_more = True
            while has_more and self.is_running:
              api_url = f"https://eg.me.logisticsbackoffice.com/api/rider-live-operations/v1/external/city/{city['live_id']}/riders?page={page}&size=100"
              try:
                res = requests.get(api_url, headers=headers, timeout=15)
                if (
                    res.status_code in [401, 403]
                    or '<!DOCTYPE html>' in res.text
                ):
                  self.log('❌ انتهت الجلسة أو طلب تسجيل دخول.')
                  session_expired = True
                  break
                if res.status_code not in [200, 201]:
                  break

                json_data = res.json()
                riders_list = (
                    json_data.get('content', json_data)
                    if isinstance(json_data, dict)
                    else json_data
                )

                if riders_list and len(riders_list) > 0:
                  for rider in riders_list:
                    raw_emp_id = str(rider.get('employee_id', '')).strip()

                    s_start = format_to_time_only_live(
                        rider.get('active_shift_started_at')
                    )
                    s_end = format_to_time_only_live(
                        rider.get('active_shift_ended_at')
                    )
                    perf = rider.get('performance', {})
                    t_spent = perf.get('time_spent', {})

                    all_riders.append({
                        'المحافظة': city['name'],
                        'الزون': (
                            rider.get('zone', {}).get('name', '')
                            if rider.get('zone')
                            else ''
                        ),
                        'ID': raw_emp_id,
                        'Name': rider.get('name', ''),
                        'Phone no.': rider.get('phone_number', ''),
                        'Current Session': (
                            f'{s_start} - {s_end}'
                            if s_start and s_end
                            else (s_start or '')
                        ),
                        'UTR': perf.get('utilization_rate', ''),
                        'Deliveries': (
                            rider.get('deliveries_info', {}).get(
                                'completed_deliveries_count', ''
                            )
                            if rider.get('deliveries_info')
                            else ''
                        ),
                        'Time worked': seconds_to_hms_live(
                            t_spent.get('worked_seconds')
                        ),
                        'Acceptance Rate': (
                            f"{int(perf['acceptance_rate'] * 100)} %"
                            if perf.get('acceptance_rate') is not None
                            else ''
                        ),
                        'Wallet Balance': (
                            rider.get('wallet_info', {}).get('balance', '')
                            if rider.get('wallet_info')
                            else ''
                        ),
                        'Breaks': t_spent.get('number_of_breaks', ''),
                        'Breaks time': seconds_to_hms_live(
                            t_spent.get('break_seconds')
                        ),
                        'Late time': seconds_to_hms_live(
                            t_spent.get('late_seconds')
                        ),
                        'Rider state': str(rider.get('status', '')).upper(),
                        'وقت السحب': current_time_live,
                    })
                  page += 1
                else:
                  has_more = False
              except Exception as e:
                self.log(
                    f"⚠️ خطأ في جلب بيانات Live لـ {city.get('en_name', '')}:"
                    f' {e}'
                )
                has_more = False

          if session_expired:
            self.log('⏳ جاري تحديث صفحة المتصفح لتحديث الجلسة...')
            self.shared_browser.refresh_page()
            time.sleep(5)
            continue

          last_live_run = time.time()
          if all_riders:
            self.latest_live_data = all_riders
            if callable(getattr(self, 'on_live_updated', None)):
              if self.log_widget:
                self.log_widget.after(0, self.on_live_updated)
            try:
              csv_path = os.path.join(SCRIPT_DIR, self.csv_filename)
              with open(csv_path, 'w', newline='', encoding='utf-8-sig') as f:
                writer = csv.DictWriter(f, fieldnames=all_riders[0].keys())
                writer.writeheader()
                writer.writerows(all_riders)

              success = self.send_data_to_sheet('live', all_riders)
              if success:
                self.log(
                    f'⚡ تم تحديث شيت LIVE بنجاح! الإجمالي: ({len(all_riders)})'
                )
              else:
                self.log('❌ تعذر رفع بيانات LIVE بعد عدة محاولات.')
            except Exception as e:
              self.log(f'❌ فشل حفظ/رفع Live: {e}')

        # 2️⃣ BOOKING CYCLE
        current_timestamp = time.time()
        time_to_booking = max(
            0,
            int(
                self.booking_interval - (current_timestamp - last_booking_run)
            ),
        )

        is_scheduled_booking = (self.mode in ['BOOKING', 'BOTH']) and (
            time_to_booking <= 0
        )
        is_manual_booking = self.manual_booking_trigger

        if (is_scheduled_booking or is_manual_booking) and self.is_running:
          self.manual_booking_trigger = False

          auth_token, cookie = self.shared_browser.get_auth_and_cookies(
              self.log
          )
          if not self.is_running or not auth_token:
            time.sleep(3)
            continue

          headers = {
              'Authorization': auth_token,
              'Cookie': cookie,
              'Accept': 'application/json',
              'User-Agent': 'Mozilla/5.0',
          }
          self.log(
              '📅 بدء دورة سحب BOOKING (Rooster) (لمدة'
              f' {self.days_ahead} أيام قادمة)...'
          )

          now_egypt = datetime.now(EGYPT_TZ)
          allowed_dates = [
              (now_egypt + timedelta(days=i)).strftime('%Y-%m-%d')
              for i in range(self.days_ahead)
          ]

          start_today = now_egypt.replace(
              hour=0, minute=0, second=0, microsecond=0
          )
          end_period = (
              start_today
              + timedelta(days=self.days_ahead)
              - timedelta(milliseconds=1)
          )

          start_utc = (
              start_today.astimezone(timezone.utc)
              .strftime('%Y-%m-%dT%H:%M:%S.000Z')
              .replace(':', '%3A')
          )
          end_utc = (
              end_period.astimezone(timezone.utc)
              .strftime('%Y-%m-%dT%H:%M:%S.999Z')
              .replace(':', '%3A')
          )

          all_bookings = []
          seen_shift_ids = set()

          for city in self.target_cities:
            if not self.is_running:
              break
            page = 0
            has_more = True
            while has_more and self.is_running:
              api_url = f"https://eg.me.logisticsbackoffice.com/api/rooster/v3/shifts?city_id={city['rooster_id']}&distinct_day_plans=true&end_at={end_utc}&page={page}&size=100&start_at={start_utc}&with_evaluations=true&with_field=id_number&with_time_zone=Africa%2FCairo"
              try:
                res = requests.get(api_url, headers=headers, timeout=15)
                if res.status_code in [401, 403]:
                  break
                data_json = res.json()
                content = data_json.get('content', [])

                if content and len(content) > 0:
                  for item in content:
                    if not isinstance(item, dict):
                      continue
                    shift_id = item.get('id', '')
                    if shift_id in seen_shift_ids:
                      continue

                    p_start_raw = item.get('start', '')
                    p_end_raw = item.get('end', '')
                    status_str = str(item.get('state', '')).upper()

                    start_dt = parse_api_datetime(p_start_raw)
                    end_dt = parse_api_datetime(p_end_raw)
                    if not start_dt:
                      continue

                    p_start_dt = start_dt.strftime('%Y-%m-%d')
                    if p_start_dt not in allowed_dates:
                      continue

                    evaluations = item.get('evaluations', [])
                    a_start_dt, a_start_tm, a_end_dt, a_end_tm, a_dur = (
                        '',
                        '',
                        '',
                        '',
                        '0:00',
                    )

                    if evaluations and len(evaluations) > 0:
                      eval_obj = evaluations[0]
                      a_start_dt_obj = parse_api_datetime(
                          eval_obj.get('start', '')
                      )
                      a_end_dt_obj = parse_api_datetime(eval_obj.get('end', ''))

                      if a_start_dt_obj and a_end_dt_obj:
                        a_start_dt = a_start_dt_obj.strftime('%Y-%m-%d')
                        a_start_tm = a_start_dt_obj.strftime('%I:%M %p')
                        a_end_dt = a_end_dt_obj.strftime('%Y-%m-%d')
                        a_end_tm = a_end_dt_obj.strftime('%I:%M %p')
                        a_dur = calculate_duration_hms(
                            eval_obj.get('start', ''), eval_obj.get('end', '')
                        )

                    seen_shift_ids.add(shift_id)
                    all_bookings.append({
                        'shift id': shift_id,
                        'employee id': item.get('employee_id', ''),
                        'employee name': item.get('employee_name', ''),
                        'starting point id': item.get('starting_point_id', ''),
                        'starting point': item.get('starting_point_name', ''),
                        'shift tag': str(item.get('tag', '')).upper(),
                        'shift status': status_str,
                        'planned start date': p_start_dt,
                        'planned end date': (
                            end_dt.strftime('%Y-%m-%d') if end_dt else ''
                        ),
                        'planned start time': start_dt.strftime('%I:%M %p'),
                        'planned end time': (
                            end_dt.strftime('%I:%M %p') if end_dt else ''
                        ),
                        'planned duration': calculate_duration_hms(
                            p_start_raw, p_end_raw
                        ),
                        'actual start date': a_start_dt,
                        'actual end date': a_end_dt,
                        'actual start time': a_start_tm,
                        'actual end time': a_end_tm,
                        'actual duration': a_dur,
                    })

                  if data_json.get('last') is True or (page + 1) >= data_json.get(
                      'total_pages', 1
                  ):
                    has_more = False
                  else:
                    page += 1
                else:
                  has_more = False
              except Exception:
                has_more = False

          last_booking_run = time.time()
          if all_bookings:
            self.latest_booking_data = all_bookings
            try:
              # تصفية الشيفتات واستبعاد الغيابات نهائياً قبل الإرسال لجوجل شيت
              booking_to_sheet = [
                  b
                  for b in all_bookings
                  if 'NO_SHOW' not in str(b.get('shift status', '')).upper()
                  and 'NO SHOW' not in str(b.get('shift status', '')).upper()
              ]
              success = self.send_data_to_sheet('booking', booking_to_sheet)
              if success:
                self.log(
                    '🚀 تم تحديث شيت BOOKING بنجاح!'
                    f' الإجمالي: ({len(booking_to_sheet)})'
                )
              else:
                self.log('❌ تعذر رفع بيانات BOOKING بعد عدة محاولات.')
            except Exception as e:
              self.log(f'❌ فشل رفع Booking: {e}')

        time.sleep(1)

      except Exception as err:
        self.log(f'⚠️ خطأ في المحرك: {err}')
        time.sleep(5)

    self.update_status('متوقف', 'black')


# ====================================================================
# 🖥️ واجهة التطبيق الرئيسية (GUI)
# ====================================================================
class AppGUI:

  def __init__(self, root):
    self.root = root
    self.root.title('لايف وحجز team mh')
    self.root.geometry('1180x880')

    self.shared_browser = SharedBrowser()
    self.shutdown_timer_id = None
    self.root.protocol('WM_DELETE_WINDOW', self.on_closing)

    self.build_shutdown_timer_ui()

    notebook = ttk.Notebook(root)
    notebook.pack(fill='both', expand=True, padx=10, pady=5)

    office2_frame = ttk.Frame(notebook)

    notebook.add(office2_frame, text=' 🏢 المكتب 2 (كافة المحافظات) ')

    office2_excluded_ids = {
        220,
        273,
        206,
        204,
        225,
        202,
        217,
        207,
        296,
        275,
    }
    office2_all_cities = [
        {'name': 'مدينة وهمية', 'live_id': 296, 'rooster_id': 296, 'en_name': 'Fake city'},
        {'name': 'بني سويف', 'live_id': 275, 'rooster_id': 275, 'en_name': 'Beni suef'},
        {'name': 'السويس', 'live_id': 227, 'rooster_id': 227, 'en_name': 'Suez'},
        {'name': 'الاسكندريه', 'live_id': 200, 'rooster_id': 200, 'en_name': 'Alexandria'},
        {'name': 'أسيوط', 'live_id': 207, 'rooster_id': 207, 'en_name': 'Assiut'},
        {'name': 'بنها', 'live_id': 216, 'rooster_id': 216, 'en_name': 'Banha'},
        {'name': 'القاهره', 'live_id': 1, 'rooster_id': 1, 'en_name': 'Cairo'},
        {'name': 'دمنهور', 'live_id': 217, 'rooster_id': 217, 'en_name': 'Damanhour'},
        {'name': 'دمياط', 'live_id': 202, 'rooster_id': 202, 'en_name': 'Damietta'},
        {'name': 'الاسماعيليه', 'live_id': 205, 'rooster_id': 205, 'en_name': 'Ismailia'},
        {'name': 'المنصوره', 'live_id': 208, 'rooster_id': 208, 'en_name': 'Mansoura'},
        {'name': 'المحله الكبرى', 'live_id': 225, 'rooster_id': 225, 'en_name': 'Al mahallah al kubra'},
        {'name': 'بورسعيد', 'live_id': 204, 'rooster_id': 204, 'en_name': 'Port said'},
        {'name': 'شبين الكوم', 'live_id': 206, 'rooster_id': 206, 'en_name': 'Shebin el koom'},
        {'name': 'الزقازيق', 'live_id': 212, 'rooster_id': 212, 'en_name': 'Zagazig'},
        {'name': 'المنيا', 'live_id': 209, 'rooster_id': 209, 'en_name': 'Minya'},
        {'name': 'الغردقه', 'live_id': 211, 'rooster_id': 211, 'en_name': 'Hurghada'},
        {'name': 'طنطا', 'live_id': 214, 'rooster_id': 214, 'en_name': 'Tanta'},
        {'name': 'سوهاج', 'live_id': 220, 'rooster_id': 220, 'en_name': 'Sohag'},
        {'name': 'كفر الشيخ', 'live_id': 273, 'rooster_id': 273, 'en_name': 'Kafr el-sheikh'},
    ]
    office2_cities = [
        city
        for city in office2_all_cities
        if city['rooster_id'] not in office2_excluded_ids
    ]

    self.office2_engine = self.build_panel(
        parent=office2_frame,
        title='لوحة التحكم - المكتب 2 (كافة المحافظات عدا المستثناة)',
        name='office2',
        cities=office2_cities,
        sheet_url='https://script.google.com/macros/s/AKfycbzL37jCFf4OvLrlVES-RN4MCOAqR_5SUIj7Os-A37xrH-nxNXR0XiTtdww2j3FURelA/exec',
        csv_file='Live_Riders_Office2.csv',
        default_live_interval=300,
        default_booking_interval=300,
        days_ahead=3,
    )

  def build_shutdown_timer_ui(self):
    frame_timer = tk.LabelFrame(
        self.root,
        text=' ⏰ أدوات عامة وأدوات المشرفين ومؤقت الإغلاق ',
        font=('Arial', 10, 'bold'),
        padx=10,
        pady=5,
    )
    frame_timer.pack(fill='x', padx=10, pady=5)

    lbl_time = tk.Label(
        frame_timer, text='وقت الإغلاق (HH:MM):', font=('Arial', 10)
    )
    lbl_time.pack(side='left', padx=5)

    self.ent_time = tk.Entry(
        frame_timer, width=8, font=('Arial', 10, 'bold'), justify='center'
    )
    self.ent_time.insert(0, '00:00')
    self.ent_time.pack(side='left', padx=5)

    self.btn_timer = tk.Button(
        frame_timer,
        text='⏳ تفعيل المؤقت',
        bg='#8b5cf6',
        fg='white',
        font=('Arial', 9, 'bold'),
        command=self.toggle_timer,
    )
    self.btn_timer.pack(side='left', padx=10)

    self.lbl_timer_status = tk.Label(
        frame_timer, text='المؤقت غير مفعل', font=('Arial', 9), fg='gray'
    )
    self.lbl_timer_status.pack(side='left', padx=10)

    self.btn_view_sup = tk.Button(
        frame_timer,
        text='👁️ عرض المشرفين المرفوعين',
        bg='#0284c7',
        fg='white',
        font=('Arial', 9, 'bold'),
        padx=8,
        command=lambda: view_supervisors_popup(self.root),
    )
    self.btn_view_sup.pack(side='right', padx=5)

    self.btn_upload_sup = tk.Button(
        frame_timer,
        text='📁 رفع شيت المشرفين (Excel/CSV)',
        bg='#059669',
        fg='white',
        font=('Arial', 9, 'bold'),
        padx=8,
        command=lambda: upload_supervisors_file(self.root),
    )
    self.btn_upload_sup.pack(side='right', padx=5)

  def toggle_timer(self):
    if self.shutdown_timer_id is None:
      target_time = self.ent_time.get().strip()
      try:
        datetime.strptime(target_time, '%H:%M')
        self.start_shutdown_timer(target_time)
        self.btn_timer.config(text='❌ إلغاء المؤقت', bg='#dc2626')
        self.lbl_timer_status.config(
            text=f'مفعل (سيتم الإغلاق الساعة {target_time})', fg='#16a34a'
        )
        messagebox.showinfo(
            'تم تفعيل المؤقت',
            f'سيتم إغلاق البرنامج والكمبيوتر تلقائياً عند الساعة {target_time}',
        )
      except ValueError:
        messagebox.showerror(
            'خطأ',
            'صيغة الوقت غير صحيحة! استخدم صيغة HH:MM مثلاً 00:00 أو 14:30',
        )
    else:
      self.cancel_shutdown_timer()

  def start_shutdown_timer(self, target_time_str):
    def check_timer():
      now_str = datetime.now(EGYPT_TZ).strftime('%H:%M')
      if now_str == target_time_str:
        self.trigger_system_shutdown()
      else:
        self.shutdown_timer_id = self.root.after(10000, check_timer)

    check_timer()

  def cancel_shutdown_timer(self):
    if self.shutdown_timer_id:
      self.root.after_cancel(self.shutdown_timer_id)
      self.shutdown_timer_id = None
      self.btn_timer.config(text='⏳ تفعيل المؤقت', bg='#8b5cf6')
      self.lbl_timer_status.config(text='المؤقت غير مفعل', fg='gray')
      messagebox.showinfo('إلغاء', 'تم إلغاء مؤقت الإغلاق بنجاح.')

  def trigger_system_shutdown(self):
    self.on_closing()
    os.system('shutdown /s /t 10')

  def build_panel(
      self,
      parent,
      title,
      name,
      cities,
      sheet_url,
      csv_file,
      default_live_interval,
      default_booking_interval,
      days_ahead,
  ):
    top_frame = tk.Frame(parent)
    top_frame.pack(fill='x', padx=10, pady=10)

    lbl_title = tk.Label(
        top_frame, text=title, font=('Arial', 14, 'bold'), anchor='w'
    )
    lbl_title.pack(side='top', fill='x', pady=5)

    status_lbl = tk.Label(
        top_frame, text='الحالة: متوقف', font=('Arial', 11), fg='black'
    )
    status_lbl.pack(side='top', anchor='w')

    time_frame = tk.LabelFrame(
        top_frame,
        text=' ⏱️ إعدادات وضع السحب والأوقات ',
        font=('Arial', 10, 'bold'),
        padx=10,
        pady=5,
    )
    time_frame.pack(side='top', fill='x', pady=8)

    tk.Label(
        time_frame, text='🎯 وضع السحب:', font=('Arial', 10, 'bold')
    ).grid(row=0, column=0, padx=5, pady=5, sticky='w')
    cb_mode = ttk.Combobox(
        time_frame,
        values=['اللايف والحجز معاً', 'لايف فقط', 'حجز فقط'],
        width=16,
        state='readonly',
        font=('Arial', 10),
    )
    cb_mode.current(0)
    cb_mode.grid(row=0, column=1, padx=5, pady=5)

    tk.Label(
        time_frame, text='تكرار LIVE (ثانية):', font=('Arial', 10)
    ).grid(row=0, column=2, padx=10, pady=5, sticky='w')
    ent_live = tk.Entry(time_frame, width=8, font=('Arial', 10, 'bold'))
    ent_live.insert(0, str(default_live_interval))
    ent_live.grid(row=0, column=3, padx=5, pady=5)

    tk.Label(
        time_frame, text='تكرار BOOKING (ثانية):', font=('Arial', 10)
    ).grid(row=0, column=4, padx=10, pady=5, sticky='w')
    ent_booking = tk.Entry(time_frame, width=8, font=('Arial', 10, 'bold'))
    ent_booking.insert(0, str(default_booking_interval))
    ent_booking.grid(row=0, column=5, padx=5, pady=5)

    tk.Label(
        time_frame, text='أيام الحجز:', font=('Arial', 10)
    ).grid(row=0, column=6, padx=10, pady=5, sticky='w')
    ent_days = tk.Entry(time_frame, width=6, font=('Arial', 10, 'bold'))
    ent_days.insert(0, str(days_ahead))
    ent_days.grid(row=0, column=7, padx=5, pady=5)

    engine = ScraperEngine(
        name=name,
        target_cities=cities,
        sheet_url=sheet_url,
        csv_filename=csv_file,
        shared_browser=self.shared_browser,
        live_interval=default_live_interval,
        booking_interval=default_booking_interval,
        days_ahead=days_ahead,
        status_label=status_lbl,
    )

    def start_with_custom_intervals():
      selected_mode = cb_mode.get()
      if selected_mode == 'لايف فقط':
        engine.mode = 'LIVE'
      elif selected_mode == 'حجز فقط':
        engine.mode = 'BOOKING'
      else:
        engine.mode = 'BOTH'

      try:
        engine.live_interval = int(ent_live.get().strip())
      except ValueError:
        engine.live_interval = default_live_interval

      try:
        engine.booking_interval = int(ent_booking.get().strip())
      except ValueError:
        engine.booking_interval = default_booking_interval

      try:
        engine.days_ahead = int(ent_days.get().strip())
      except ValueError:
        engine.days_ahead = days_ahead

      engine.log(
          f'⚙️ نمط السحب: [{selected_mode}] | LIVE كل {engine.live_interval} ث'
          f' | BOOKING كل {engine.booking_interval} ث | سحب {engine.days_ahead}'
          ' أيام'
      )
      engine.start()

    btn_frame = tk.Frame(top_frame)
    btn_frame.pack(side='top', fill='x', pady=5)

    btn_start = tk.Button(
        btn_frame,
        text='▶️ تشغيل',
        bg='#2ec4b6',
        fg='white',
        font=('Arial', 11, 'bold'),
        width=10,
        command=start_with_custom_intervals,
    )
    btn_start.pack(side='left', padx=5)

    btn_stop = tk.Button(
        btn_frame,
        text='⏹️ إيقاف',
        bg='#e71d36',
        fg='white',
        font=('Arial', 11, 'bold'),
        width=10,
        command=engine.stop,
    )
    btn_stop.pack(side='left', padx=5)

    btn_manual_live = tk.Button(
        btn_frame,
        text='🔄 تحديث اللايف الآن',
        bg='#059669',
        fg='white',
        font=('Arial', 10, 'bold'),
        padx=10,
        command=engine.trigger_manual_live,
    )
    btn_manual_live.pack(side='left', padx=8)

    btn_manual_booking = tk.Button(
        btn_frame,
        text='🔄 تحديث الحجز الآن',
        bg='#0284c7',
        fg='white',
        font=('Arial', 10, 'bold'),
        padx=10,
        command=engine.trigger_manual_booking,
    )
    btn_manual_booking.pack(side='left', padx=8)

    status_box = tk.LabelFrame(
        top_frame,
        text=' 🔍 عرض قائمة الطيارين والنسب حسب الحالة (Live) ',
        font=('Arial', 10, 'bold'),
        padx=10,
        pady=5,
    )
    status_box.pack(side='top', fill='x', pady=8)

    chart_canvas = tk.Canvas(
        status_box,
        width=130,
        height=130,
        bg=status_box.cget('bg'),
        highlightthickness=0,
    )
    chart_canvas.grid(row=0, column=6, rowspan=3, padx=15, pady=5)

    def draw_donut_chart(canvas, total, working_count):
      canvas.delete('all')

      pct = (working_count / total * 100) if total > 0 else 0.0
      angle = (working_count / total * 360) if total > 0 else 0

      x0, y0, x1, y1 = 10, 10, 120, 120

      canvas.create_oval(x0, y0, x1, y1, fill='#e2e8f0', outline='')

      if angle > 0:
        canvas.create_arc(
            x0, y0, x1, y1, start=90, extent=-angle, fill='#10b981', outline=''
        )

      canvas.create_oval(28, 28, 102, 102, fill='#ffffff', outline='')

      canvas.create_text(
          65, 48, text=f'{pct:.1f}%', font=('Arial', 11, 'bold'), fill='#0f172a'
      )
      canvas.create_text(
          65,
          68,
          text=f'{working_count} / {total}',
          font=('Arial', 8, 'bold'),
          fill='#059669',
      )
      canvas.create_text(
          65, 84, text='شغال حالياً', font=('Arial', 8), fill='#64748b'
      )

    lbl_summary_pct = tk.Label(
        status_box,
        text='📊 النسب العامة: بانتظار سحب البيانات...',
        font=('Arial', 9, 'bold'),
        fg='#1e293b',
        anchor='w',
    )
    lbl_summary_pct.grid(
        row=0, column=0, columnspan=6, sticky='w', padx=5, pady=(0, 5)
    )

    status_buttons_info = [
        ('WORKING', 'working', '#10b981'),
        ('BREAK', 'break', '#f59e0b'),
        ('LATE', 'late', '#ef4444'),
        ('STARTING', 'starting', '#3b82f6'),
        ('READY', 'ready', '#06b6d4'),
        ('ENDING', 'ending', '#6b7280'),
    ]

    status_btn_widgets = {}

    def update_percentages_ui():
      if not engine.latest_live_data:
        return

      total_live = len(engine.latest_live_data)
      if total_live == 0:
        return

      working_count = sum(
          1
          for r in engine.latest_live_data
          if 'WORKING' in str(r.get('Rider state', '')).upper()
          and 'NOT_WORKING' not in str(r.get('Rider state', '')).upper()
      )

      draw_donut_chart(chart_canvas, total_live, working_count)

      summary_parts = []

      for disp_name, key, _ in status_buttons_info:
        count = sum(
            1
            for r in engine.latest_live_data
            if key.upper() in str(r.get('Rider state', '')).upper()
            and (
                key.upper() != 'WORKING'
                or 'NOT_WORKING' not in str(r.get('Rider state', '')).upper()
            )
        )
        pct = (count / total_live) * 100

        if key in status_btn_widgets:
          status_btn_widgets[key].config(
              text=f'{disp_name}\n{count} طيار ({pct:.1f}%)'
          )

        if count > 0:
          summary_parts.append(f'{disp_name}: {pct:.1f}% ({count})')

      full_summary_str = (
          f'📊 الإجمالي: ({total_live} طيار) | الشغال:'
          f' {working_count} ({((working_count/total_live)*100):.1f}%) | '
          + ' | '.join(summary_parts)
      )
      lbl_summary_pct.config(text=full_summary_str)

    engine.on_live_updated = update_percentages_ui

    def open_status_popup(status_key, display_name):
      if not engine.latest_live_data:
        engine.log(
            f'⚠️ لا توجد بيانات لايف مسحوبة بعد لمُعاينة حالة'
            f' ({display_name}).'
        )
        return

      if status_key.upper() == 'WORKING':
        status_filtered = [
            r
            for r in engine.latest_live_data
            if 'WORKING' in str(r.get('Rider state', '')).upper()
            and 'NOT_WORKING' not in str(r.get('Rider state', '')).upper()
        ]
      else:
        status_filtered = [
            r
            for r in engine.latest_live_data
            if status_key.upper() in str(r.get('Rider state', '')).upper()
        ]

      total_live = len(engine.latest_live_data)
      state_pct = (
          (len(status_filtered) / total_live * 100) if total_live > 0 else 0
      )

      tot_worked_sec = 0
      for r in status_filtered:
        tw = str(r.get('Time worked', '0:00:00'))
        try:
          p = list(map(int, tw.split(':')))
          if len(p) == 3:
            tot_worked_sec += p[0] * 3600 + p[1] * 60 + p[2]
          elif len(p) == 2:
            tot_worked_sec += p[0] * 60 + p[1]
        except Exception:
          pass

      w_hrs = tot_worked_sec // 3600
      w_mins = (tot_worked_sec % 3600) // 60

      today_str = datetime.now(EGYPT_TZ).strftime('%Y-%m-%d')
      tot_booked_mins = 0
      if engine.latest_booking_data:
        rider_booking_map = {}
        for b in engine.latest_booking_data:
          if b.get('planned start date') == today_str:
            emp_id = str(b.get('employee id', '')).strip()
            if not emp_id:
              continue
            dur_str = b.get('planned duration', '0:00')
            try:
              parts = dur_str.split(':')
              mins = int(parts[0]) * 60 + int(parts[1])
            except Exception:
              mins = 0
            rider_booking_map[emp_id] = rider_booking_map.get(emp_id, 0) + mins

        for r in status_filtered:
          emp_id = str(r.get('ID', '')).strip()
          tot_booked_mins += rider_booking_map.get(emp_id, 0)

      b_hrs = tot_booked_mins // 60
      b_mins = tot_booked_mins % 60

      popup = tk.Toplevel(parent)
      popup.title(f'قائمة الطيارين - حالة: {display_name} ({state_pct:.1f}%)')
      popup.geometry('1280x580')

      top_popup_frame = tk.Frame(popup, pady=6, padx=10)
      top_popup_frame.pack(side='top', fill='x')

      info_text = (
          f'عدد الطيارين: {len(status_filtered)} من أصل {total_live}'
          f' ({state_pct:.1f}%)  |  ساعات العمل الفعلية:'
          f' {w_hrs}:{w_mins:02d} ساعة  |  ساعات الحجز (اليوم):'
          f' {b_hrs}:{b_mins:02d} ساعة'
      )

      lbl_info = tk.Label(
          top_popup_frame,
          text=info_text,
          font=('Arial', 11, 'bold'),
          fg='#1e293b',
      )
      lbl_info.pack(side='left', padx=5)

      action_bar = tk.Frame(popup, pady=4, padx=10, bg='#f1f5f9')
      action_bar.pack(side='top', fill='x', padx=10, pady=2)

      tk.Label(
          action_bar,
          text='🔍 فلتر وتصفية القائمة:',
          font=('Arial', 10, 'bold'),
          bg='#f1f5f9',
      ).pack(side='left', padx=5)
      ent_status_search = tk.Entry(action_bar, width=22, font=('Arial', 10))
      ent_status_search.pack(side='left', padx=5)

      cols = (
          'ID',
          'Name',
          'اسم المشرف',
          'Phone no.',
          'المحافظة',
          'الزون',
          'Current Session',
          'Deliveries',
          'Time worked',
          'Late time',
          'Rider state',
      )
      tree = ttk.Treeview(popup, columns=cols, show='headings', height=15)

      for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=100, anchor='center')

      tree.column('Name', width=140, anchor='w')
      tree.column('اسم المشرف', width=130, anchor='center')
      tree.column('Current Session', width=110)

      vsb = ttk.Scrollbar(popup, orient='vertical', command=tree.yview)
      hsb = ttk.Scrollbar(popup, orient='horizontal', command=tree.xview)
      tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

      vsb.pack(side='right', fill='y')
      hsb.pack(side='bottom', fill='x')
      tree.pack(fill='both', expand=True, padx=10, pady=5)

      def refresh_popup_table(*args):
        for item in tree.get_children():
          tree.delete(item)
        query = ent_status_search.get().strip().lower()

        for r in status_filtered:
          emp_id = str(r.get('ID', '')).strip()
          emp_name = str(r.get('Name', '')).lower()
          emp_zone = str(r.get('الزون', '')).lower()
          emp_phone = str(r.get('Phone no.', '')).lower()

          supervisor_name = supervisors_map.get(emp_id, 'غير محدد')

          if query and not (
              query in emp_id.lower()
              or query in emp_name
              or query in emp_zone
              or query in emp_phone
              or query in supervisor_name.lower()
          ):
            continue

          tree.insert(
              '',
              'end',
              values=(
                  emp_id,
                  r.get('Name', ''),
                  supervisor_name,
                  r.get('Phone no.', ''),
                  r.get('المحافظة', ''),
                  r.get('الزون', ''),
                  r.get('Current Session', ''),
                  r.get('Deliveries', ''),
                  r.get('Time worked', ''),
                  r.get('Late time', ''),
                  r.get('Rider state', ''),
              ),
          )

      ent_status_search.bind('<KeyRelease>', refresh_popup_table)

      btn_export = tk.Button(
          action_bar,
          text='📊 تنزيل إكسيل',
          bg='#16a34a',
          fg='white',
          font=('Arial', 9, 'bold'),
          command=lambda: export_treeview_to_csv(
              tree, f'Riders_{display_name}.csv'
          ),
      )
      btn_export.pack(side='right', padx=4)

      btn_copy = tk.Button(
          action_bar,
          text='📋 نسخ الجدول',
          bg='#0284c7',
          fg='white',
          font=('Arial', 9, 'bold'),
          command=lambda: copy_treeview_to_clipboard(popup, tree),
      )
      btn_copy.pack(side='right', padx=4)

      btn_snap = tk.Button(
          action_bar,
          text='📸 اسكرين شوت',
          bg='#d97706',
          fg='white',
          font=('Arial', 9, 'bold'),
          command=lambda: take_window_screenshot(
              tree, f'Snapshot_{display_name}'
          ),
      )
      btn_snap.pack(side='right', padx=4)

      refresh_popup_table()

    # ====================================================================
    # 🚨 نافذة الغيابات المجمعة (بدون تكرار المناديب مع حساب الشيفتات والساعات)
    # ====================================================================
    def open_absences_popup():
      if not engine.latest_booking_data:
        messagebox.showinfo(
            'تنبيه',
            'لا توجد بيانات حجز مسحوبة بعد. يرجى الضغط على زر "🔄 تحديث الحجز'
            ' الآن" أولاً.',
        )
        return

      today_str = datetime.now(EGYPT_TZ).strftime('%Y-%m-%d')

      # تجميع غيابات اليوم حسب ID المندوب
      aggregated_absences = {}
      total_absent_shifts_count = 0
      total_absent_minutes = 0

      for b in engine.latest_booking_data:
        p_date = b.get('planned start date')
        status = str(b.get('shift status', '')).upper()

        if p_date == today_str and (
            'NO_SHOW' in status or 'NO SHOW' in status
        ):
          emp_id = str(b.get('employee id', '')).strip()
          if not emp_id:
            continue

          emp_name = str(b.get('employee name', '')).strip()
          sp_name = str(b.get('starting point', '')).strip()
          dur_str = b.get('planned duration', '0:00')

          # حساب دقائق الشيفت الملغي
          try:
            parts = str(dur_str).split(':')
            mins = int(parts[0]) * 60 + int(parts[1])
          except Exception:
            mins = 0

          s_time = b.get('planned start time', '')
          e_time = b.get('planned end time', '')
          shift_time_range = (
              f'{s_time} - {e_time}' if (s_time and e_time) else ''
          )

          total_absent_shifts_count += 1
          total_absent_minutes += mins

          if emp_id not in aggregated_absences:
            aggregated_absences[emp_id] = {
                'emp_id': emp_id,
                'emp_name': emp_name,
                'starting_points': set(),
                'shift_count': 0,
                'total_minutes': 0,
                'shift_times': [],
            }

          if sp_name:
            aggregated_absences[emp_id]['starting_points'].add(sp_name)
          aggregated_absences[emp_id]['shift_count'] += 1
          aggregated_absences[emp_id]['total_minutes'] += mins
          if shift_time_range:
            aggregated_absences[emp_id]['shift_times'].append(shift_time_range)

      popup = tk.Toplevel(parent)
      popup.title(f'❌ غيابات اليوم المجمعة ({today_str})')
      popup.geometry('1250x580')

      tot_h = total_absent_minutes // 60
      tot_m = total_absent_minutes % 60

      top_bar = tk.Frame(popup, pady=6, padx=10)
      top_bar.pack(side='top', fill='x')

      lbl_title_abs = tk.Label(
          top_bar,
          text=(
              f'❌ إجمالي المناديب الغائبين: {len(aggregated_absences)} طيار  | '
              f' إجمالي الشيفتات الملغاة: {total_absent_shifts_count}  | '
              f' إجمالي ساعات الغياب: {tot_h}:{tot_m:02d} ساعة'
          ),
          font=('Arial', 11, 'bold'),
          fg='#dc2626',
      )
      lbl_title_abs.pack(side='left', padx=5)

      action_bar = tk.Frame(popup, pady=4, padx=10, bg='#f1f5f9')
      action_bar.pack(side='top', fill='x', padx=10, pady=2)

      tk.Label(
          action_bar,
          text='🔍 بحث (ID / اسم / مشرف / منطقة):',
          font=('Arial', 10, 'bold'),
          bg='#f1f5f9',
      ).pack(side='left', padx=5)
      ent_abs_search = tk.Entry(action_bar, width=25, font=('Arial', 10))
      ent_abs_search.pack(side='left', padx=5)

      cols = (
          'ID المندوب',
          'اسم المندوب',
          'اسم المشرف',
          'النقاط / الزونات',
          'عدد شيفتات الغياب',
          'إجمالي ساعات الغياب',
          'أوقات الشيفتات الملغاة',
      )
      tree = ttk.Treeview(popup, columns=cols, show='headings', height=16)

      for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor='center')

      tree.column('اسم المندوب', width=160, anchor='w')
      tree.column('اسم المشرف', width=140, anchor='center')
      tree.column('النقاط / الزونات', width=160, anchor='center')
      tree.column('أوقات الشيفتات الملغاة', width=240, anchor='center')

      vsb = ttk.Scrollbar(popup, orient='vertical', command=tree.yview)
      hsb = ttk.Scrollbar(popup, orient='horizontal', command=tree.xview)
      tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

      vsb.pack(side='right', fill='y')
      hsb.pack(side='bottom', fill='x')
      tree.pack(fill='both', expand=True, padx=10, pady=5)

      def refresh_absences_table(*args):
        for item in tree.get_children():
          tree.delete(item)

        q = ent_abs_search.get().strip().lower()

        for emp_id, data in aggregated_absences.items():
          emp_name = data['emp_name']
          sup_name = supervisors_map.get(emp_id, 'غير محدد')
          sp_str = (
              ' - '.join(data['starting_points'])
              if data['starting_points']
              else 'غير محدد'
          )

          h = data['total_minutes'] // 60
          m = data['total_minutes'] % 60
          dur_formatted = f'{h}:{m:02d} ساعة'
          times_formatted = ' | '.join(data['shift_times'])

          if q and not (
              q in emp_id.lower()
              or q in emp_name.lower()
              or q in sp_str.lower()
              or q in sup_name.lower()
          ):
            continue

          tree.insert(
              '',
              'end',
              values=(
                  emp_id,
                  emp_name,
                  sup_name,
                  sp_str,
                  data['shift_count'],
                  dur_formatted,
                  times_formatted,
              ),
          )

      ent_abs_search.bind('<KeyRelease>', refresh_absences_table)

      btn_export = tk.Button(
          action_bar,
          text='📊 تنزيل إكسيل',
          bg='#16a34a',
          fg='white',
          font=('Arial', 9, 'bold'),
          command=lambda: export_treeview_to_csv(
              tree, f'Absences_Summary_{today_str}.csv'
          ),
      )
      btn_export.pack(side='right', padx=4)

      btn_copy = tk.Button(
          action_bar,
          text='📋 نسخ الجدول',
          bg='#0284c7',
          fg='white',
          font=('Arial', 9, 'bold'),
          command=lambda: copy_treeview_to_clipboard(popup, tree),
      )
      btn_copy.pack(side='right', padx=4)

      btn_snap = tk.Button(
          action_bar,
          text='📸 اسكرين شوت',
          bg='#d97706',
          fg='white',
          font=('Arial', 9, 'bold'),
          command=lambda: take_window_screenshot(
              tree, f'Snapshot_Absences_{today_str}'
          ),
      )
      btn_snap.pack(side='right', padx=4)

      refresh_absences_table()

    def open_extension_filter_popup():
      if not engine.latest_live_data:
        engine.log('⚠️ لا توجد بيانات لايف مسحوبة بعد.')
        return

      today_str = datetime.now(EGYPT_TZ).strftime('%Y-%m-%d')

      rider_booking_stats = {}
      if engine.latest_booking_data:
        for b in engine.latest_booking_data:
          if b.get('planned start date') == today_str:
            emp_id = str(b.get('employee id', '')).strip()
            if not emp_id:
              continue
            dur_str = b.get('planned duration', '0:00')
            try:
              parts = dur_str.split(':')
              mins = int(parts[0]) * 60 + int(parts[1])
            except Exception:
              mins = 0

            if emp_id not in rider_booking_stats:
              rider_booking_stats[emp_id] = {
                  'shift_count': 0,
                  'total_minutes': 0,
              }
            rider_booking_stats[emp_id]['shift_count'] += 1
            rider_booking_stats[emp_id]['total_minutes'] += mins

      popup = tk.Toplevel(parent)
      popup.title('🛠️ فلتر التمديد واستعلام شيفتات المناديب المحجوزة')
      popup.geometry('1280x660')

      filter_frame = tk.LabelFrame(
          popup,
          text=' ⚙️ خيارات تصفية المناديب للاتصال أو التمديد ',
          font=('Arial', 10, 'bold'),
          padx=10,
          pady=5,
      )
      filter_frame.pack(side='top', fill='x', padx=10, pady=5)

      tk.Label(filter_frame, text='حالة المندوب:', font=('Arial', 10)).grid(
          row=0, column=0, padx=5, pady=5, sticky='w'
      )
      cb_state = ttk.Combobox(
          filter_frame,
          values=[
              'WORKING (شغال حالياً)',
              'الكل (المناديب الفاتحة)',
              'BREAK',
          ],
          width=20,
          state='readonly',
      )
      cb_state.current(0)
      cb_state.grid(row=0, column=1, padx=5, pady=5)

      tk.Label(
          filter_frame,
          text='حاجز ساعات اليوم أقل من أو يساوي:',
          font=('Arial', 10),
      ).grid(row=0, column=2, padx=5, pady=5, sticky='w')
      ent_max_booked = tk.Entry(
          filter_frame, width=8, font=('Arial', 10, 'bold')
      )
      ent_max_booked.grid(row=0, column=3, padx=5, pady=5)

      tk.Label(
          filter_frame,
          text='شغال عمل فعلي أكثر من أو يساوي (ساعة):',
          font=('Arial', 10),
      ).grid(row=0, column=4, padx=5, pady=5, sticky='w')
      ent_min_worked = tk.Entry(
          filter_frame, width=8, font=('Arial', 10, 'bold')
      )
      ent_min_worked.grid(row=0, column=5, padx=5, pady=5)

      tk.Label(
          filter_frame, text='بحث خاص (ID / اسم / مشرف):', font=('Arial', 10)
      ).grid(row=1, column=0, padx=5, pady=5, sticky='w')
      ent_search = tk.Entry(filter_frame, width=20, font=('Arial', 10))
      ent_search.grid(row=1, column=1, padx=5, pady=5)

      action_tools_frame = tk.Frame(popup, padx=10, pady=2)
      action_tools_frame.pack(side='top', fill='x')

      lbl_stats = tk.Label(
          action_tools_frame,
          text='',
          font=('Arial', 11, 'bold'),
          fg='#1e293b',
      )
      lbl_stats.pack(side='left', padx=5)

      cols = (
          'ID',
          'Name',
          'اسم المشرف',
          'Phone no.',
          'المحافظة',
          'الزون',
          'Rider state',
          'Time worked',
          'عدد الشيفتات',
          'إجمالي الساعات المحجوزة',
          'Current Session',
      )
      tree = ttk.Treeview(popup, columns=cols, show='headings', height=16)

      for col in cols:
        tree.heading(col, text=col)
        tree.column(col, width=105, anchor='center')

      tree.column('Name', width=140, anchor='w')
      tree.column('اسم المشرف', width=120, anchor='center')
      tree.column('Current Session', width=110)

      vsb = ttk.Scrollbar(popup, orient='vertical', command=tree.yview)
      hsb = ttk.Scrollbar(popup, orient='horizontal', command=tree.xview)
      tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

      vsb.pack(side='right', fill='y')
      hsb.pack(side='bottom', fill='x')
      tree.pack(fill='both', expand=True, padx=10, pady=5)

      btn_ext_export = tk.Button(
          action_tools_frame,
          text='📊 تنزيل إكسيل',
          bg='#16a34a',
          fg='white',
          font=('Arial', 9, 'bold'),
          command=lambda: export_treeview_to_csv(tree, 'Extension_Riders.csv'),
      )
      btn_ext_export.pack(side='right', padx=4)

      btn_ext_copy = tk.Button(
          action_tools_frame,
          text='📋 نسخ الجدول',
          bg='#0284c7',
          fg='white',
          font=('Arial', 9, 'bold'),
          command=lambda: copy_treeview_to_clipboard(popup, tree),
      )
      btn_ext_copy.pack(side='right', padx=4)

      btn_ext_snap = tk.Button(
          action_tools_frame,
          text='📸 اسكرين شوت',
          bg='#d97706',
          fg='white',
          font=('Arial', 9, 'bold'),
          command=lambda: take_window_screenshot(
              tree, 'Snapshot_Extension_Filter'
          ),
      )
      btn_ext_snap.pack(side='right', padx=4)

      def apply_filter():
        for item in tree.get_children():
          tree.delete(item)

        st_val = cb_state.get()
        max_b_val = ent_max_booked.get().strip()
        min_w_val = ent_min_worked.get().strip()
        search_q = ent_search.get().strip().lower()

        try:
          max_b_hrs = float(max_b_val) if max_b_val else None
        except ValueError:
          max_b_hrs = None

        try:
          min_w_hrs = float(min_w_val) if min_w_val else None
        except ValueError:
          min_w_hrs = None

        match_count = 0
        tot_booked_mins = 0

        for live_r in engine.latest_live_data:
          emp_id = str(live_r.get('ID', '')).strip()
          emp_name = str(live_r.get('Name', '')).strip()
          r_state = str(live_r.get('Rider state', '')).upper()

          supervisor_name = supervisors_map.get(emp_id, 'غير محدد')

          if st_val == 'WORKING (شغال حالياً)':
            if 'WORKING' not in r_state or 'NOT_WORKING' in r_state:
              continue
          elif st_val == 'BREAK':
            if 'BREAK' not in r_state:
              continue
          else:
            is_working = 'WORKING' in r_state and 'NOT_WORKING' not in r_state
            is_break = 'BREAK' in r_state
            if not (is_working or is_break):
              continue

          if search_q:
            if (
                search_q not in emp_id.lower()
                and search_q not in emp_name.lower()
                and search_q not in supervisor_name.lower()
            ):
              continue

          b_stats = rider_booking_stats.get(
              emp_id, {'shift_count': 0, 'total_minutes': 0}
          )
          b_mins = b_stats['total_minutes']
          b_hrs_dec = b_mins / 60.0

          if max_b_hrs is not None and b_hrs_dec > max_b_hrs:
            continue

          tw = str(live_r.get('Time worked', '0:00:00'))
          w_mins = 0
          try:
            p = list(map(int, tw.split(':')))
            if len(p) == 3:
              w_mins = p[0] * 60 + p[1]
            elif len(p) == 2:
              w_mins = p[0] * 60 + p[1]
          except Exception:
            w_mins = 0
          w_hrs_dec = w_mins / 60.0

          if min_w_hrs is not None and w_hrs_dec < min_w_hrs:
            continue

          match_count += 1
          tot_booked_mins += b_mins

          b_h = b_mins // 60
          b_m = b_mins % 60
          b_str = f'{b_h}:{b_m:02d} ساعة'

          tree.insert(
              '',
              'end',
              values=(
                  emp_id,
                  emp_name,
                  supervisor_name,
                  live_r.get('Phone no.', ''),
                  live_r.get('المحافظة', ''),
                  live_r.get('الزون', ''),
                  live_r.get('Rider state', ''),
                  live_r.get('Time worked', ''),
                  b_stats['shift_count'],
                  b_str,
                  live_r.get('Current Session', ''),
              ),
          )

        tot_h = tot_booked_mins // 60
        tot_m = tot_booked_mins % 60
        lbl_stats.config(
            text=(
                f'📊 عدد المناديب المطابقين للشرط: {match_count} مندوب  | '
                ' إجمالي الساعات المحجوزة للنتائج:'
                f' {tot_h}:{tot_m:02d} ساعة'
            )
        )

      btn_apply = tk.Button(
          filter_frame,
          text='🔍 تطبيق الفلتر والبحث',
          bg='#2563eb',
          fg='white',
          font=('Arial', 10, 'bold'),
          padx=10,
          command=apply_filter,
      )
      btn_apply.grid(row=1, column=4, columnspan=2, padx=5, pady=5, sticky='e')

      apply_filter()

    for idx, (disp_name, key, col) in enumerate(status_buttons_info):
      btn = tk.Button(
          status_box,
          text=disp_name,
          bg=col,
          fg='white',
          font=('Arial', 9, 'bold'),
          padx=6,
          pady=3,
          command=lambda k=key, d=disp_name: open_status_popup(k, d),
      )
      btn.grid(row=1, column=idx, padx=4, pady=5)
      status_btn_widgets[key] = btn

    bottom_tools_frame = tk.Frame(status_box)
    bottom_tools_frame.grid(
        row=2, column=0, columnspan=6, padx=4, pady=5, sticky='w'
    )

    btn_extension_filter = tk.Button(
        bottom_tools_frame,
        text='🛠️ فلتر التمديد والمناديب الشغالة (حساب الساعات والشيفتات)',
        bg='#0f766e',
        fg='white',
        font=('Arial', 10, 'bold'),
        padx=10,
        pady=4,
        command=open_extension_filter_popup,
    )
    btn_extension_filter.pack(side='left', padx=(0, 10))

    btn_absences = tk.Button(
        bottom_tools_frame,
        text='❌ غيابات اليوم',
        bg='#dc2626',
        fg='white',
        font=('Arial', 10, 'bold'),
        padx=12,
        pady=4,
        command=open_absences_popup,
    )
    btn_absences.pack(side='left')

    log_area = scrolledtext.ScrolledText(
        parent, wrap=tk.WORD, font=('Consolas', 10), state='disabled'
    )
    log_area.pack(fill='both', expand=True, padx=10, pady=10)

    engine.log_widget = log_area
    return engine

  def on_closing(self):
    if self.shutdown_timer_id:
      self.root.after_cancel(self.shutdown_timer_id)
      self.shutdown_timer_id = None

    if hasattr(self, 'alex_engine'):
      self.alex_engine.stop()
    if hasattr(self, 'hurghada_engine'):
      self.hurghada_engine.stop()
    if hasattr(self, 'office2_engine'):
      self.office2_engine.stop()

    self.shared_browser.close()
    self.root.destroy()


if __name__ == '__main__':
  root = tk.Tk()
  app = AppGUI(root)
  root.mainloop()