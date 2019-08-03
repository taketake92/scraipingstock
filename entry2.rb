require "google_drive"
require 'nokogiri'
require 'open-uri'
require 'date'

session = GoogleDrive::Session.from_config("config.json")

mCodes = [
  2269,
3382,
3990,
4502,
4755,
5108,
5713,
6504,
8316,
9064,
9437,
9984,

]

a = 0


mCodes.each{|mCode|

# https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxx/
# 事前に書き込みたいスプレッドシートを作成し、上記スプレッドシートのURL(「xxx」の部分)を以下のように指定する
sheets = session.spreadsheet_by_key("1F3EzCs0gDOsizeTqqBWwh94YBipEb0-d5-6lktuH01s").worksheets[a]



startRow = 5 #行

  url = "https://kabutan.jp/stock/?code=#{mCode}"

  p url

  charset = nil

  html = open(url) do |f|
      charset = f.charset
      f.read
  end

  doc = Nokogiri::HTML.parse(html, nil, charset)

  a0 =
  p a0

  a1 = doc.xpath("//*[@id='kobetsu_left']/table[1]/tbody/tr[1]/td[1]").inner_text
  p a1

  a2 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[2]/td[1]').inner_text
  p a2

  a3 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[2]/td[3]').inner_text.delete!("(,)")
  p a3

  a4 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[3]/td[1]').inner_text
  p a4

  a5 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[3]/td[3]').inner_text.delete!("(,)")
  p a5

  a6 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[4]/td[1]').inner_text
  p a6

  a7 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[1]/td').inner_text
  p a7

  a8 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[2]/td').inner_text
  p a8

  a9 = doc.xpath('//*[@id="stockinfo_i1"]/div[2]/dl/dd[1]/span').inner_text
  p a9

  a10 = doc.xpath('//*[@id="stockinfo_i1"]/div[2]/dl/dd[2]/span').inner_text
  p a10

  a11 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[3]/td').inner_text
  p a11

  a12 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[7]/td').inner_text
  p a12


  # スプレッドシートへの書き込み
  # sheets[startRow,1] = Date.today.strftime("%Y/%m/%d") #日付
  sheets[startRow,1] = Date.today #始値
  sheets[startRow,2] = a1 #始値
  sheets[startRow,3] = a2 #高値
  sheets[startRow,4] = a3 #高値(時刻)
  sheets[startRow,5] = a4 #安値
  sheets[startRow,6] = a5 #安値(時刻)
  sheets[startRow,7] = a6 #終値
  sheets[startRow,8] = a7 #出来高
  sheets[startRow,9] = a8 #売買代金
  sheets[startRow,10] = a9 #前日比(数値)
  sheets[startRow,11] = a10 #前日比(数値)
  sheets[startRow,12] = a11 #VWAP
  sheets[startRow,13] = a12 #時価総額

  a += 1

# シートの保存
sheets.save

}
