require "google_drive"
require 'nokogiri'
require 'open-uri'
require 'date'
require 'CSV'

session = GoogleDrive::Session.from_config("config.json")

mCodes = [
'0000',
2002,
2269,
2282,
2501,
2502,
2503,
2531,
2801,
2802,
2871,
2914,
3101,
3103,
3401,
3402,
3861,
3863,
3405,
3407,
4004,
4005,
4021,
4042,
4043,
4061,
4063,
4183,
4188,
4208,
4272,
4452,
4631,
4901,
4911,
4151,
4502,
4503,
4506,
4507,
4519,
4523,
4568,
4578,
5019,
5020,
5101,
5108,
5201,
5202,
5214,
5232,
5233,
5301,
5332,
5333,
5401,
5406,
5411,
5541,
3436,
5703,
5706,
5707,
5711,
5713,
5714,
5801,
5802,
5803,
5901,
5631,
6103,
6113,
6301,
6302,
6305,
6326,
6361,
6367,
6471,
6472,
6473,
7004,
7011,
7013,
3105,
6479,
6501,
6503,
6504,
6506,
6645,
6674,
6701,
6702,
6703,
6724,
6752,
6758,
6762,
6770,
6841,
6857,
6902,
6952,
6954,
6971,
6976,
6988,
7735,
7751,
7752,
8035,
7003,
7012,
7201,
7202,
7203,
7205,
7211,
7261,
7267,
7269,
7270,
7272,
4543,
4902,
7731,
7733,
7762,
7832,
7911,
7912,
7951,
1332,
1333,
1605,
1721,
1801,
1802,
1803,
1808,
1812,
1925,
1928,
1963,
2768,
8001,
8002,
8015,
8031,
8053,
8058,
3086,
3099,
3382,
8028,
8233,
8252,
8267,
9983,
7186,
8303,
8304,
8306,
8308,
8309,
8316,
8331,
8354,
8355,
8411,
8601,
8604,
8628,
8630,
8725,
8729,
8750,
8766,
8795,
8253,
3289,
8801,
8802,
8804,
8830,
9001,
9005,
9007,
9008,
9009,
9020,
9021,
9022,
9062,
9064,
9101,
9104,
9107,
9202,
9301,
9412,
9432,
9433,
9437,
9613,
9984,
9501,
9502,
9503,
9531,
9532,
2432,
4324,
4689,
4704,
4751,
4755,
6098,
6178,
9602,
9681,
9735,
9766,
3990,
9984,
2002,
3556,
3665,
3987,
4320,
4427,
4726,
5704,
6184,
6754,
6963,
7974,
9467,
3110,
3038,
2326,
2384,
3046,
3064,
3349,
3479,
3687,
6920,
7309,
7550,
7564,
7816,
9934,
1357,
4588,
4344,
7748,
3906,
3356,
9434,
6541,
3092,
3967,
6861,
3627,
1570,
6033,
6071,
3652,
3962,
4238,
6155,
6246,

]

a = 0


mCodes.each{|mCode|

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
  a1 = a1.size == 3 ? a1 : a1.delete!(",")
  p a1

  a2 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[2]/td[1]').inner_text
  a2 = a2.size == 3 ? a2 : a2.delete!(",")
  p a2

  a3 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[2]/td[3]').inner_text.delete!("(,)")
  p a3

  a4 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[3]/td[1]').inner_text
  a4 = a4.size == 3 ? a4 : a4.delete!(",")
  p a4

  a5 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[3]/td[3]').inner_text.delete!("(,)")
  p a5

  a6 = doc.xpath('//*[@id="kobetsu_left"]/table[1]/tbody/tr[4]/td[1]').inner_text
  a6 = a6.size == 3 ? a6 : a6.delete!(",")
  p a6

  a7 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[1]/td').inner_text.gsub("\u00A0", "")
  a7 = a7.include?(",") ? a7.delete!(",") : a7
  p a7

  a8 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[2]/td').inner_text.gsub("\u00A0", "")
  a8 = a8.include?(",") ? a8.delete!(",") : a8
  p a8

  a9 = doc.xpath('//*[@id="stockinfo_i1"]/div[2]/dl/dd[1]/span').inner_text
  p a9

  a10 = doc.xpath('//*[@id="stockinfo_i1"]/div[2]/dl/dd[2]/span').inner_text
  p a10

  a11 = doc.xpath('//*[@id="kobetsu_left"]/table[2]/tbody/tr[3]/td').inner_text.gsub("\u00A0", "")
  a11 = a11.include?(",") ? a11.delete!(",") : a11

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



  test = CSV.open("#{mCode}.csv","w") do |csv|
    ##翌日wをaに変えて上書き#

  csv << [
          mCode.to_s,
          Date.today.strftime("%Y-%m-%d"),
          a1,
          a2,
          a3,
          a4,
          a5,
          a6,
          a7,
          a8,
          a9,
          a10,
          a11,
          a12,
        ]

end

sleep(3)

}
