from telethon.sync import TelegramClient
import pandas as pd
import time
from telethon.tl.types import ReactionEmoji, ReactionPaid, ReactionCustomEmoji, MessageMediaPhoto
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from datetime import datetime, timedelta

POSITIVE_EMOJIS = {'👍', '❤', '🔥', '😊', '😂', '🥰', '👏', '⚡', '❤‍🔥', '🫡', '🤗', '😍', '👌', '😁', '💯', '🙏', '🤩'}
NEGATIVE_EMOJIS = {'🤡', '👎', '💩', '🥱'}

api_id = "25492203"
api_hash = "0f1f1ab020df21383b4699bea789871a"

def process_reactions(reactions):
    positive = []
    negative = []
    neutral = []
    paid = []
    total = 0
    if reactions:
        for reaction in reactions.results:
            if isinstance(reaction.reaction, ReactionEmoji):
                emoticon = reaction.reaction.emoticon
                if emoticon in POSITIVE_EMOJIS:
                    positive.append(f"{emoticon} {reaction.count}")
                elif emoticon in NEGATIVE_EMOJIS:
                    negative.append(f"{emoticon} {reaction.count}")
                else:
                    neutral.append(f"{emoticon} {reaction.count}")
            elif isinstance(reaction.reaction, ReactionPaid):
                paid.append(f"{reaction.count}")
            elif isinstance(reaction.reaction, ReactionCustomEmoji):
                neutral.append(f"? {reaction.count}")
        total = sum(reaction.count for reaction in reactions.results)
    return positive, negative, neutral, paid, total

def calculate_engagement_rate(views, reactions, comments, forwards, paid_reactions):
    if views > 0:
        total_engagement = reactions + comments + forwards + paid_reactions
        return round(total_engagement / views * 100, 2)
    return 0

def save_to_excel(df, filename, channel_username, start_date, end_date):
    summary_row = df[df['Дата'] == 'Итого']
    average_row = df[df['Дата'] == 'Среднее']
    df = df[~df['Дата'].isin(['Итого', 'Среднее'])]
    df = pd.concat([summary_row, average_row, df], ignore_index=True)
    df.to_excel(filename, index=False)
    wb = load_workbook(filename)
    ws = wb.active
    info_text = f"Данные из канала {channel_username} (https://t.me/{channel_username}) за период с {start_date.strftime('%Y-%m-%d %H:%M:%S')} по {end_date.strftime('%Y-%m-%d %H:%M:%S')}"
    ws.insert_rows(1)
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=len(df.columns))
    ws.cell(row=1, column=1, value=info_text)
    fill = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    font = Font(bold=True, size=12)
    alignment = Alignment(horizontal="center", vertical="center")
    for cell in ws[1]:
        cell.fill = fill
        cell.font = font
        cell.alignment = alignment
    column_widths = {
        'A': 10, 'B': 20, 'C': 30, 'D': 8, 'E': 15, 'F': 13, 'G': 11,
        'H': 17, 'I': 11, 'J': 17, 'K': 11, 'L': 20, 'M': 11, 'N': 13,
        'O': 14, 'P': 15, 'Q': 7,
    }
    for col, width in column_widths.items():
        ws.column_dimensions[col].width = width
    header_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    for cell in ws[2]:
        cell.fill = header_fill
    summary_fill = PatternFill(start_color="66FF66", end_color="66FF66", fill_type="solid")
    for cell in ws[3]:
        cell.fill = summary_fill
    for cell in ws[4]:
        cell.fill = summary_fill
    ws.freeze_panes = 'A5'
    wb.save(filename)

def run_parser(channel_username, start_date, end_date):
    with TelegramClient('session_name', api_id, api_hash) as client:
        messages = []
        albums = {}
        for message in client.iter_messages(channel_username):
            message_date_naive = message.date.replace(tzinfo=None)
            if message_date_naive < start_date:
                break
            if message_date_naive > end_date:
                continue
            grouped_id = getattr(message, 'grouped_id', None)
            if grouped_id:
                if grouped_id not in albums:
                    albums[grouped_id] = []
                albums[grouped_id].append(message)
            else:
                post_type = 'Опрос' if message.poll else 'Текст' if message.text else 'Фото' if isinstance(message.media, MessageMediaPhoto) else 'Другое'
                message_link = f"https://t.me/{channel_username}/{message.id}" if channel_username else 'Нет ссылки'
                positive, negative, neutral, paid, total_reactions = process_reactions(message.reactions)
                views = message.views or 0
                comments = message.replies.replies if message.replies else 0
                forwards = message.forwards or 0
                paid_reactions = sum(int(p) for p in paid)
                er_percentage = calculate_engagement_rate(views, total_reactions, comments, forwards, paid_reactions)
                messages.append({
                    'Тип': post_type,
                    'Дата': message_date_naive,
                    'Текст': message.text or '',
                    'Длина': len(message.text or ''),
                    'Ссылка': message_link,
                    'Просмотры': views,
                    'Реакции': total_reactions,
                    'Позитивные': ' '.join(positive),
                    'Всего (+)': sum(int(r.split(' ')[1]) for r in positive),
                    'Негативные': ' '.join(negative),
                    'Всего (-)': sum(int(r.split(' ')[1]) for r in negative),
                    'Неопределённые': ' '.join(neutral),
                    'Всего': sum(int(r.split(' ')[1]) for r in neutral),
                    'Платные': ' '.join(paid),
                    'Комменты': comments,
                    'Пересылки': forwards,
                    'ER%': er_percentage
                })
            time.sleep(0.1)

        for grouped_id, album_messages in albums.items():
            first_message = album_messages[0]
            views = first_message.views or 0
            positive = []
            negative = []
            neutral = []
            paid = []
            total_reactions = 0
            paid_reactions_total = 0
            for message in album_messages:
                p, n, neu, p_paid, total = process_reactions(message.reactions)
                positive.extend(p)
                negative.extend(n)
                neutral.extend(neu)
                paid.extend(p_paid)
                total_reactions += total
                paid_reactions_total += sum(int(p) for p in p_paid)
            comments = sum(m.replies.replies if m.replies else 0 for m in album_messages)
            forwards = sum(m.forwards or 0 for m in album_messages)
            er_percentage = calculate_engagement_rate(views, total_reactions, comments, forwards, paid_reactions_total)
            messages.append({
                'Тип': 'Альбом',
                'Дата': first_message.date.replace(tzinfo=None),
                'Текст': ' '.join(m.text or '' for m in album_messages),
                'Длина': sum(len(m.text or '') for m in album_messages),
                'Ссылка': f"https://t.me/{channel_username}/{first_message.id}" if channel_username else 'Нет ссылки',
                'Просмотры': views,
                'Реакции': total_reactions,
                'Позитивные': ' '.join(positive),
                'Всего (+)': sum(int(r.split(' ')[1]) for r in positive),
                'Негативные': ' '.join(negative),
                'Всего (-)': sum(int(r.split(' ')[1]) for r in negative),
                'Неопределённые': ' '.join(neutral),
                'Всего': sum(int(r.split(' ')[1]) for r in neutral),
                'Платные': ' '.join(paid),
                'Комменты': comments,
                'Пересылки': forwards,
                'ER%': er_percentage
            })

        df = pd.DataFrame(messages)
        df = df.sort_values(by='Дата', ascending=False).reset_index(drop=True)
        numeric_columns = ['Просмотры', 'Реакции', 'Всего (+)', 'Всего (-)', 'Всего', 'Платные', 'Комменты', 'Пересылки', 'ER%']
        for col in numeric_columns:
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
        summary = {
            'Тип': '', 'Дата': 'Итого', 'Текст': f'Количество постов: {len(df)}', 'Длина': df['Длина'].sum(),
            'Ссылка': '', 'Просмотры': df['Просмотры'].sum(), 'Реакции': df['Реакции'].sum(), 'Позитивные': '',
            'Всего (+)': df['Всего (+)'].sum(), 'Негативные': '', 'Всего (-)': df['Всего (-)'].sum(),
            'Неопределённые': '', 'Всего': df['Всего'].sum(), 'Платные': df['Платные'].sum(),
            'Комменты': df['Комменты'].sum(), 'Пересылки': df['Пересылки'].sum(), 'ER%': ''
        }
        average = {
            'Тип': '', 'Дата': 'Среднее', 'Текст': '', 'Длина': round(df[df['Длина'] > 0]['Длина'].mean(), 2),
            'Ссылка': '', 'Просмотры': round(df['Просмотры'].mean(), 2), 'Реакции': round(df['Реакции'].mean(), 2),
            'Позитивные': '', 'Всего (+)': round(df['Всего (+)'].mean(), 2), 'Негативные': '',
            'Всего (-)': round(df['Всего (-)'].mean(), 2), 'Неопределённые': '', 'Всего': round(df['Всего'].mean(), 2),
            'Платные': round(df['Платные'].mean(), 2), 'Комменты': round(df['Комменты'].mean(), 2),
            'Пересылки': round(df['Пересылки'].mean(), 2), 'ER%': round(df['ER%'].mean(), 2)
        }
        df = pd.concat([pd.DataFrame([summary]), pd.DataFrame([average]), df], ignore_index=True)
        filename = f"{channel_username}_{start_date.strftime('%Y-%m-%d')}-{end_date.strftime('%Y-%m-%d')}.xlsx"
        save_to_excel(df, filename, channel_username, start_date, end_date)
        return filename