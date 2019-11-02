require "google_drive"
require 'nokogiri'
require 'open-uri'
require 'date'
require 'CSV'

#https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxx/
#事前に書き込みたいスプレッドシートを作成し、上記スプレッドシートのURL(「xxx」の部分)を以下のように指定する
# sheets = session.spreadsheet_by_key(id).worksheets[0]

# p sheets

mCodes = [
'0000',
2002,
2269,

]

mCodes.each{|mCode|

a = 0

session = GoogleDrive::Session.from_config("config.json")

id = "1cQbEb0RtlTE1aSn21qwJ_fzyrX5D_aRojtXXENMAopM"

#テンプレートからコピーして新規にsheet作成
#テンプレート読み込み
templeate_sheets = ""
copy_sheets = ""
templeate_sheets = session.spreadsheet_by_key(id)
#テンプレート作成
p mCode
copy_sheets = templeate_sheets.copy("#{mCode}test")
#テンプレートのid取得
sheet_id = copy_sheets.id
#テンプレートでシートインスタンス
sheets = session.spreadsheet_by_key(sheet_id).worksheets[0]


  # https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxx/
  # 事前に書き込みたいスプレッドシートを作成し、上記スプレッドシートのURL(「xxx」の部分)を以下のように指定する
  # sheets = session.spreadsheet_by_key("1F3EzCs0gDOsizeTqqBWwh94YBipEb0-d5-6lktuH01s").worksheets[a]



  startRow = 5 #行

  url = "https://kabutan.jp/stock/?code=#{mCode}"

  p url

  charset = nil

  html = open(url) do |f|
      charset = f.charset
      f.read
  end

  doc = Nokogiri::HTML.parse(html, nil, charset)

  #始値
  a1 = doc.xpath("//*[@id='kobetsu_left']/table[1]/tbody/tr[1]/td[1]").inner_text
  a1 = a1.include?(",") ? a1.delete!(",") : a1
  p a1

  a2 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[2]/td[1]').inner_text
  a2 = a2.include?(",") ? a2.delete!(",") : a2
  p a2

  a3 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[2]/td[3]').inner_text.delete!("(,)")
  p a3

  a4 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[3]/td[1]').inner_text
  a4 = a4.include?(",") ? a4.delete!(",") : a4
  p a4

  a5 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[3]/td[3]').inner_text.delete!("(,)")
  p a5

  a6 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[4]/td[1]').inner_text
  a6 = a6.include?(",") ? a6.delete!(",") : a6
  p a6

  a7 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[1]/td').inner_text.gsub("\u00A0", "")
  a7 = a7.include?(",") ? a7.delete!(",") : a7
  p a7

  a8 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[2]/td').inner_text.gsub("\u00A0", "")
  a8 = a8.include?(",") ? a8.delete!(",") : a8
  p a8

  a9 = doc.xpath('//*[@id="stockinfo_i1"]/div[2]/dl/dd[1]/span').inner_text
  a9 = a9.empty? ? "0" : a9.include?(",") ? a9.delete!(",") : a9
  p a9

  a10 = doc.xpath('//*[@id="stockinfo_i1"]/div[2]/dl/dd[2]/span').inner_text
  a10 = a10.empty? ? "0" : a10.include?(",") ? a10.delete!(",") : a10
  p a10

  a11 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[3]/td').inner_text.gsub("\u00A0", "")
  a11 = a11.include?(",") ? a11.delete!(",") : a11
  p a11

  a12 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[4]/td').inner_text.gsub("\u00A0", "")
  a12 = a12.include?(",") ? a12.delete!(",") : a12
  p a12

  a13 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[7]/td').inner_text.gsub("\u00A0", "")
  a13 = a13.include?(",") ? a13.delete!(",") : a13
  p a13


  # スプレッドシートへの書き込み
  # # sheets[startRow,1] = Date.today.strftime("%Y/%m/%d") #日付
  # sheets[startRow,1] = "2019/08/02" #始値
  # sheets[startRow,2] = a1 #始値
  # sheets[startRow,3] = a2 #高値
  # sheets[startRow,4] = a3 #高値(時刻)
  # sheets[startRow,5] = a4 #安値
  # sheets[startRow,6] = a5 #安値(時刻)
  # sheets[startRow,7] = a6 #終値
  # sheets[startRow,8] = a7 #出来高
  # sheets[startRow,9] = a8 #売買代金
  # sheets[startRow,10] = a9 #前日比(数値)
  # sheets[startRow,11] = a10 #前日比(数値)
  # sheets[startRow,12] = a11 #VWAP
  # sheets[startRow,13] = a12 #約定回数
  # sheets[startRow,14] = a13 #時価総額


  a += 1

  # シートの保存
  # sheets.save

  #https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxx/
  #事前に書き込みたいスプレッドシートを作成し、上記スプレッドシートのURL(「xxx」の部分)を以下のように指定する
  sheets = session.spreadsheet_by_key(id).worksheets[0]

  aaa = CSV.open("./#{mCode}.csv").count + 2


  #スプレッドシートへの書き込み
  sheets[aaa,1] = mCode.to_s #銘柄コード
  sheets[aaa,2] = "2019-10-18" #日付
  sheets[aaa,3] = a9 #前日比
  sheets[aaa,4] = a10 #前日比(㌫)
  sheets[aaa,5] = a3 #始値
  sheets[aaa,7] = a4 #高値
  sheets[aaa,8] = a5 #高値(時刻)
  sheets[aaa,9] = a6 #安値
  sheets[aaa,10] = a7 #安値(時刻)
  sheets[aaa,11] = a8 #終値
  sheets[aaa,12] = a9 #出来高
  sheets[aaa,13] = a13 #売買代金
  sheets[aaa,14] = a12 #約定回数
  sheets[aaa,15] = a13 #時価総額
  
  #シートの保存
  sheets.save

  test = CSV.open("#{mCode}.csv","a") do |csv|
    ##aデータ更新#
    ##w新規でテーブル作成

  # csv << [
  #         mCode.to_s,
  #         # Date.today.strftime("%Y-%m-%d"),
  #         "2019-10-18",
  #         a1,
  #         a2,
  #         a3,
  #         a4,
  #         a5,
  #         a6,
  #         a7,
  #         a8,
  #         a9,
  #         a10,
  #         a11,
  #         a12,
  #         a13,
  #       ]
  # 
  end

  c = rand(5..6)

  sleep(c)

}
