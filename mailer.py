import pandas as pd
import smtplib
import time
import random
import re
import json
import os
from datetime import date
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email import encoders
from email.header import Header
import argparse
import ssl
from pathlib import Path


# ==============================
# 1. КЛАСС РАБОТЫ С EMAIL-ДАННЫМИ
# ==============================
class EmailDataProcessor:
    def __init__(self, file_path, sheet_name=0, email_column_hint='email'):
        self.file_path = file_path
        self.sheet_name = sheet_name
        self.email_column_hint = email_column_hint.lower()
        self.df = self._read_file()
        self._clean_and_extract()

    def _read_file(self):
        if self.file_path.lower().endswith('.csv'):
            df = pd.read_csv(self.file_path, dtype=str, keep_default_na=False)
        else:
            df = pd.read_excel(self.file_path, sheet_name=self.sheet_name, dtype=str, keep_default_na=False)
        df.columns = [str(c) for c in df.columns]
        return df

    def _find_email_column(self):
        for col in self.df.columns:
            if self.email_column_hint in col.lower():
                return col
        for col in self.df.columns:
            if col.strip().lower() == 'email':
                return col
        raise ValueError("❌ Не найден столбец с email. Проверь название или укажи email_column_hint.")

    EMAIL_RE = re.compile(
        r'([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})',
        re.IGNORECASE
    )

    def _extract_emails_from_text(self, text):
        if not isinstance(text, str) or not text:
            return []

        found = self.EMAIL_RE.findall(text)
        cleaned = []
        for e in found:
            e = e.strip().lower()
            if e not in cleaned:
                cleaned.append(e)
        return cleaned[:3]  # максимум 3 email из одной ячейки

    def _clean_and_extract(self):
        email_col = self._find_email_column()
        rows = []

        for idx, row in self.df.iterrows():
            cell = row.get(email_col, '')
            emails = self._extract_emails_from_text(cell)
            if not emails:
                continue

            for em in emails:
                new_row = row.to_dict()
                new_row['_email'] = em
                rows.append(new_row)

        self.emails_df = pd.DataFrame(rows, columns=[*self.df.columns, '_email']).drop_duplicates(subset=['_email'])
        print(f"✅ Найдено уникальных email: {len(self.emails_df)}")

    def get_emails_list(self):
        return self.emails_df['_email'].tolist()

    def filter_mail_ru_emails(self):
        """Фильтрует только почты mail.ru"""
        mail_ru_emails = []
        for email in self.emails_df['_email']:
            if email.endswith('@mail.ru') or email.endswith('@inbox.ru') or email.endswith('@list.ru') or email.endswith('@bk.ru'):
                mail_ru_emails.append(email)
        
        print(f"📧 Найдено почт mail.ru: {len(mail_ru_emails)}")
        return mail_ru_emails


# ==============================
# 2. КЛАСС УМНОЙ РАССЫЛКИ
# ==============================
class SmartEmailSender:
    def __init__(self, email_configs, state_file='email_state.json'):
        self.email_configs = email_configs
        self.state_file = state_file
        self.load_state()

    def load_state(self):
        if os.path.exists(self.state_file):
            with open(self.state_file, 'r', encoding='utf-8') as f:
                self.state = json.load(f)
            if self.state.get('last_date') != str(date.today()):
                print("🔄 Новый день, сбрасываем лимиты...")
                self.reset_daily_limits()
        else:
            self.state = {
                'last_date': str(date.today()),
                'sent_today': 0,
                'emails_sent': {},
                'processed_emails': [],
                'current_email_index': 0
            }
            self.save_state()

    def save_state(self):
        with open(self.state_file, 'w', encoding='utf-8') as f:
            json.dump(self.state, f, ensure_ascii=False, indent=2)

    def reset_daily_limits(self):
        self.state['last_date'] = str(date.today())
        self.state['sent_today'] = 0
        self.state['emails_sent'] = {}
        self.save_state()

    def get_available_email_config(self):
        total = len(self.email_configs)
        for attempt in range(total):
            config = self.email_configs[self.state['current_email_index']]
            self.state['current_email_index'] = (self.state['current_email_index'] + 1) % total
            email = config['email']
            sent = self.state['emails_sent'].get(email, 0)
            if sent < config['daily_limit']:
                return config
        return None

    def update_sent_count(self, email):
        self.state['emails_sent'][email] = self.state['emails_sent'].get(email, 0) + 1
        self.state['sent_today'] += 1
        self.save_state()

    def get_smart_delay(self):
        # Пауза между отправками
        base = random.choice([2, 3, 4, 5])
        variation = random.uniform(-0.5, 1)
        return max(1, base + variation)

    def create_antispam_headers(self, msg, to_email, from_email):
        msg['Reply-To'] = from_email
        msg['Return-Path'] = from_email
        # Supply a real unsubscribe endpoint in your own private configuration.
        msg['Precedence'] = 'bulk'
        msg['X-Auto-Response-Suppress'] = 'OOF, AutoReply'
        return msg

    def attach_images_to_email(self, msg):
        """Прикрепляет оба изображения и возвращает статус"""
        images_attached = []
        
        try:
            # Первое изображение - сертификат
            if os.path.exists("inline_image.png"):
                with open("inline_image.png", "rb") as img_file:
                    img_data = img_file.read()
                
                image1 = MIMEImage(img_data, name="inline_image.png")
                image1.add_header('Content-ID', '<inline_image>')
                image1.add_header('Content-Disposition', 'inline; filename="inline_image.png"')
                msg.attach(image1)
                images_attached.append('inline_image')
            else:
                print("   ⚠️ Файл inline_image.png не найден")

            # Второе изображение - image001
            if os.path.exists("image001.png"):
                with open("image001.png", "rb") as img_file:
                    img_data = img_file.read()
                
                image2 = MIMEImage(img_data, name="image001.png")
                image2.add_header('Content-ID', '<image001>')
                image2.add_header('Content-Disposition', 'inline; filename="image001.png"')
                msg.attach(image2)
                images_attached.append('image001')
            else:
                print("   ⚠️ Файл image001.png не найден")

            return images_attached
            
        except Exception as e:
            print(f"   ❌ Ошибка прикрепления изображений: {e}")
            return []

    def attach_files(self, msg):
        try:
            files = [
                ("attachment.docx", "attachment.docx"),
            ]
            for path, name in files:
                if os.path.exists(path):
                    with open(path, "rb") as f:
                        if path.endswith(".docx"):
                            part = MIMEBase('application', 'vnd.openxmlformats-officedocument.wordprocessingml.document')
                            part.set_payload(f.read())
                            encoders.encode_base64(part)
                        elif path.endswith(".pdf"):
                            part = MIMEBase('application', 'pdf')
                            part.set_payload(f.read())
                            encoders.encode_base64(part)
                        
                        part.add_header('Content-Disposition', f'attachment; filename="{name}"')
                        msg.attach(part)
                else:
                    print(f"   ⚠️ Файл не найден: {path}")
        except Exception as e:
            print(f"   ❌ Ошибка вложений: {e}")

    def send_single_email(self, config, to_email, subject, html_body, text_body):
        try:
            server = smtplib.SMTP(config['smtp_server'], config['smtp_port'], timeout=30)
            server.starttls(context=ssl.create_default_context())
            server.login(config['email'], config['password'])

            msg = MIMEMultipart('related')
            msg['From'] = config['email']
            msg['To'] = to_email
            msg['Subject'] = Header(subject, 'utf-8')

            msg = self.create_antispam_headers(msg, to_email, config['email'])

            # Прикрепляем оба изображения
            images_attached = self.attach_images_to_email(msg)

            # Создаем HTML с изображениями
            html_body_with_images = html_body

            # Заменяем плейсхолдеры на изображения (БЕЗ ПОДПИСЕЙ)
            if 'inline_image' in images_attached:
                html_body_with_images = html_body_with_images.replace(
                    '<!--CERTIFICATE_IMAGE-->',
                    '''
                    <div style="text-align: center; margin: 20px 0;">
                        <img src="cid:inline_image" 
                             alt="Иллюстрация" 
                             style="max-width: 100%; height: auto; border: 1px solid #ddd; border-radius: 5px; max-width: 600px;">
                    </div>
                    '''
                )
            else:
                html_body_with_images = html_body_with_images.replace(
                    '<!--CERTIFICATE_IMAGE-->',
                    ''
                )

            if 'image001' in images_attached:
                html_body_with_images = html_body_with_images.replace(
                    '<!--IMAGE001-->',
                    '''
                    <div style="margin: 15px 0;">
                        <img src="cid:image001" 
                             alt="Иллюстрация" 
                             style="max-width: 100%; height: auto; border: 1px solid #ddd; border-radius: 5px;">
                    </div>
                    '''
                )
            else:
                html_body_with_images = html_body_with_images.replace(
                    '<!--IMAGE001-->',
                    ''
                )

            # Создаем альтернативную часть с HTML и текстом
            alt = MIMEMultipart('alternative')
            alt.attach(MIMEText(text_body, 'plain', 'utf-8'))
            alt.attach(MIMEText(html_body_with_images, 'html', 'utf-8'))
            msg.attach(alt)

            # Прикрепляем остальные файлы
            self.attach_files(msg)

            server.sendmail(config['email'], to_email, msg.as_string())
            server.quit()

            self.update_sent_count(config['email'])
            return True
        except Exception as e:
            print(f"❌ Ошибка отправки ({config['email']} → {to_email}): {e}")
            return False

    def calculate_remaining_capacity(self):
        total_capacity = sum(c['daily_limit'] for c in self.email_configs)
        return total_capacity - self.state['sent_today']

    def send_bulk_emails(self, emails_list, resume=True):
        if resume:
            processed = set(self.state['processed_emails'])
            to_send = [e for e in emails_list if e not in processed]
        else:
            to_send = emails_list
            self.state['processed_emails'] = []
            self.save_state()

        subject, html_body, text_body = create_email()

        print(f"🚀 Начинаем рассылку ({len(to_send)} получателей)")
        print(f"📧 Осталось писем на сегодня: {self.calculate_remaining_capacity()}")

        sent = 0
        start_time = time.time()
        
        for i, email_to in enumerate(to_send, 1):
            if self.calculate_remaining_capacity() <= 0:
                print("⚠️ Достигнут дневной лимит.")
                break

            config = self.get_available_email_config()
            if not config:
                print("⚠️ Все аккаунты исчерпали лимит.")
                break

            ok = self.send_single_email(config, email_to, subject, html_body, text_body)
            if ok:
                sent += 1
                self.state['processed_emails'].append(email_to)
                self.save_state()
                # Упрощенный вывод для скорости
                if i % 10 == 0:
                    elapsed = time.time() - start_time
                    speed = i / elapsed if elapsed > 0 else 0
                    print(f"{i}/{len(to_send)} ✅ {sent} отправлено ({speed:.1f} писем/сек)")
            else:
                if i % 10 == 0:
                    print(f"{i}/{len(to_send)} ❌ Ошибки")

            # Пауза между отправками
            delay = self.get_smart_delay()
            time.sleep(delay)

            # Периодическая пауза
            if i % 100 == 0:
                pause = random.randint(15, 30)
                print(f"⏸️ Короткий перерыв {pause} сек...")
                time.sleep(pause)

        total_time = time.time() - start_time
        speed = sent / total_time if total_time > 0 else 0
        print(f"\n📊 Отправлено писем: {sent}")
        print(f"⏱️ Общее время: {total_time:.1f} сек")
        print(f"🚀 Средняя скорость: {speed:.1f} писем/сек")


def create_email():
    """Read private content; no subject or signature is shipped in the repository."""
    content_dir = Path(os.environ.get('MAIL_CONTENT_DIR', 'private'))
    subject = (content_dir / 'subject.txt').read_text(encoding='utf-8').strip()
    html_body = (content_dir / 'body.html').read_text(encoding='utf-8')
    text_body = (content_dir / 'body.txt').read_text(encoding='utf-8')
    if not subject or any(c in subject for c in ('\r', '\n')):
        raise ValueError('Тема должна быть непустой и занимать одну строку.')
    return subject, html_body, text_body


def load_accounts():
    """Secrets are resolved from environment variables, never from committed files."""
    path = Path(os.environ.get('MAIL_ACCOUNTS_FILE', 'private/accounts.json'))
    accounts = json.loads(path.read_text(encoding='utf-8'))
    if not isinstance(accounts, list) or not accounts:
        raise ValueError('Нужен непустой список SMTP-аккаунтов.')
    for account in accounts:
        account['password'] = os.environ[account.pop('password_env')]
        if int(account['daily_limit']) <= 0:
            raise ValueError('Дневной лимит должен быть положительным.')
    return accounts


def main():
    parser = argparse.ArgumentParser(description='CSV / Excel email automation')
    parser.add_argument('--file', default='examples/recipients.csv')
    parser.add_argument('--domain', help='Optional exact recipient domain')
    parser.add_argument('--send', action='store_true', help='Enable actual SMTP delivery')
    args = parser.parse_args()
    processor = EmailDataProcessor(args.file)
    recipients = processor.get_emails_list()
    if args.domain:
        recipients = [email for email in recipients if email.rsplit('@', 1)[1] == args.domain.lower()]
    if not args.send:
        print(f'Предпросмотр: {len(recipients)} уникальных адресов. Отправка отключена.')
        return
    if not recipients:
        print('Нет адресатов для отправки.')
        return
    accounts = load_accounts()
    create_email()  # Validate content before opening an SMTP connection.
    Path('private').mkdir(exist_ok=True)
    sender = SmartEmailSender(accounts, state_file='private/email_state.json')
    sender.send_bulk_emails(recipients, resume=True)


if __name__ == '__main__':
    main()
