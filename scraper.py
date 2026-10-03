from datetime import datetime
import os
import pandas as pd
import pytz
import requests

# Устанавливаем бакинское время
baku_tz = pytz.timezone("Asia/Baku")
now_baku = datetime.now(baku_tz)

date_str = now_baku.strftime("%Y-%m-%d")
time_str = now_baku.strftime("%H:%M")

url = "https://c2b-fbusiness.customs.gov.az/api/v1/transports/count"
rows_to_save = []

try:
  response = requests.get(url, timeout=15)
  if response.status_code == 200:
    data_json = response.json()

    # Проходим по странам в списке data
    countries = data_json.get("data", [])
    for country in countries:
      transports = country.get("transports", [])
      for post in transports:
        post_name = post.get("postName", "Неизвестный пост")
        cars_in_queue = post.get("countInQueue", 0)
        cars_in_direction = post.get("countInWay", 0)

        rows_to_save.append({
            "Дата": date_str,
            "Время": time_str,
            "Таможенный пост": post_name,
            "Машины в очереди": cars_in_queue,
            "Машины в направлении": cars_in_direction,
            "Статус": "Успешно",
        })
  else:
    print(f"Ошибка сервера: {response.status_code}")

except Exception as e:
  print(f"Ошибка при запросе к API: {e}")

# Если список пуст по какой-то причине
if not rows_to_save:
  rows_to_save.append({
      "Дата": date_str,
      "Время": time_str,
      "Таможенный пост": "Ошибка данных",
      "Машины в очереди": 0,
      "Машины в направлении": 0,
      "Статус": "Ошибка",
  })

df_new = pd.DataFrame(rows_to_save)
excel_file = "queue_data.xlsx"

# Накапливаем историю в Excel
if os.path.exists(excel_file):
  df_existing = pd.read_excel(excel_file)
  df_combined = pd.concat([df_existing, df_new], ignore_index=True)
else:
  df_combined = df_new

df_combined.to_excel(excel_file, index=False)
print("Статистика по всем постам успешно сохранена в Excel!") # trigger 
