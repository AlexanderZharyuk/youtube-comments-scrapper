import pyyoutube
from dotenv import load_dotenv
import os
import csv
from openpyxl import Workbook


def get_channel_id(api, channel_username):
    channel_info = api.search(q=channel_username, return_json=True)['items'][0]
    channel_id = channel_info['id']['channelId']
    return channel_id


def get_videos_ids(api, channel_id):
    """Данный метод выдает 100 последних видео и их ID"""
    channel_activity = api.get_activities_by_channel(channel_id=channel_id, return_json=True, count=None)
    videos_ids = [video['contentDetails']['upload']['videoId'] for video in channel_activity['items']]
    return videos_ids


def check_comment(string, find_word):
    if string.lower().find(find_word) == -1:
        return False
    return True


def get_video_link(api, video_id):
    player_data = api.get_video_by_id(video_id=video_id).to_dict()['items'][0]['player']['embedHtml']
    start_index = player_data.find('/')
    end_index = player_data[start_index:].find('"') + start_index
    video_link = f"https://{player_data[start_index + 2: end_index].replace('embed', 'watch')}"
    return video_link


def add_to_csv_data(row_info):
    with open('comments_data.csv', 'a', newline='') as csvfile:
        writer = csv.writer(csvfile)
        writer.writerow(row_info)


def generated_xlsx(csv_file):
    wb = Workbook()
    ws = wb.active

    with open(csv_file, 'r') as csv_file:
        for row in csv.reader(csv_file):
            ws.append(row)

    wb.save('matched_comments.xlsx')


def get_video_comments(api, video_id, check_word):
    comments_on_video = api.get_comment_threads(video_id=video_id, count=None, order='relevance', return_json=True)
    for comment in comments_on_video['items']:
        get_comment_text = comment['snippet']['topLevelComment']['snippet']['textDisplay']

        if check_comment(get_comment_text, check_word):
            title = connected_api.get_video_by_id(video_id=video_id).to_dict()['items'][0]['snippet']['title']
            url = get_video_link(api, video_id)
            video_info = [title, get_comment_text, url]
            add_to_csv_data(video_info)


if __name__ == '__main__':
    load_dotenv()
    connected_api = pyyoutube.Api(api_key=os.getenv('API_KEY'))
    channel_name = 'Диджитализируй'
    channel_id = get_channel_id(connected_api, channel_name)
    channel_videos_ids = get_videos_ids(connected_api, channel_id)
    find_word = 'ментор'
    for video_id in channel_videos_ids:
        get_video_comments(connected_api, video_id, find_word)

    generated_xlsx('comments_data.csv')
    os.remove('comments_data.csv')
