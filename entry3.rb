require "google_drive"
require 'nokogiri'
require 'open-uri'

session = GoogleDrive::Session.from_config("config.json")

# https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxx/
# 事前に書き込みたいスプレッドシートを作成し、上記スプレッドシートのURL(「xxx」の部分)を以下のように指定する
sheets = session.spreadsheet_by_key("11VvXRFIbtc8EFcv0gG6VQd7AE6SQf8761nQJEK8wAJU").worksheets[0]

mCodes=[

  # 1447,
  # 1449,
  # 1730,
  # 2176,
  # 2334,
  # 2493,
  # 3042,
  # 3135,
  # 3195,
  # 3461,
  # 3550,
  # 3556,
  # 3674,
  # 3695,
  # 3814,
  # 3823,
  # 3849,
  # 3910,
  # 3931,
  # 3933,
  # 3970,
  # 3995,
  # 3999,
  # 4240,
  # 4308,
  # 4320,
  # 4335,
  # 4381,
  # 4394,
  # 4395,
  # 4396,
  # 4437,
  # 4442,
  # 4728,
  # 4770,
  # 5704,
  # 6070,
  # 6093,
  # 6096,
  # 6173,
  # 6192,
  # 6391,
  # 6537,
  # 6551,
  # 6558,
  # 6563,
  # 6575,
  # 6577,
  # 6579,
  # 6709,
  # 6897,
  # 7044,
  # 7063,
  # 7064,
  # 7066,
  # 7162,
  # 7320,
  # 7559,
  # 7674,
  # 9266,
  # 9271,
  # 9285,
  # 9325,
]

startRow = 2

mCodes.each{|mCode|
  url = "https://minkabu.jp/stock/#{mCode}"

  p url

  charset = nil

  html = open(url) do |f|
      charset = f.charset
      f.read
  end

  doc = Nokogiri::HTML.parse(html, nil, charset)

  doc = Nokogiri::HTML.parse(html, nil, charset)

  a0 = doc.xpath("//*[@id='sh_field_body']/div/div[1]/div/div/div[2]/div/div[1]/div/div[2]/div[2]/table").text
  b = a0[137...156]
  d = b.index("倍")
  e = b[0,d]

  p e




#   # スプレッドシートへの書き込み
#   sheets[startRow,2] = a0 #企業名
#   sheets[startRow,3] = a1 #始値
#   sheets[startRow,4] = a2 #高値
#   sheets[startRow,5] = a3 #高値(時刻)
#   sheets[startRow,6] = a4 #安値
#   sheets[startRow,7] = a5 #安値(時刻)
#   sheets[startRow,8] = a6 #終値
#   sheets[startRow,9] = a7 #出来高
#   sheets[startRow,10] = a8 #売買代金
#   sheets[startRow,11] = a9 #前日比(数値)
#   sheets[startRow,12] = a10 #前日比(数値)
#
#   startRow += 1
#
# # シートの保存
# sheets.save

# sleep(2)
}
