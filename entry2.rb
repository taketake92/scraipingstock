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
#ここからスクリーニング結果
7827,
3477,
8903,
3467,
3604,
5697,
1795,
6580,
8885,
1844,
3695,
8886,
6335,
6694,
6867,
3241,
3524,
6267,
6570,
3750,
3814,
9625,
2974,
6757,
3461,
1739,
3294,
7781,
3647,
2436,
3489,
6032,
4356,
1431,
2599,
1826,
5941,
4952,
1994,
2162,
4323,
8893,
3440,
6362,
1420,
6164,
3538,
8144,
4098,
3550,
3323,
2153,
7559,
8732,
3934,
5922,
3919,
6159,
7722,
8887,
6403,
6059,
7059,
2676,
9233,
9057,
7539,
1960,
4987,
3452,
6331,
1726,
7745,
5290,
3482,
8929,
2301,
6535,
5013,
6089,
6677,
8917,
3772,
9270,
3546,
6257,
7199,
8877,
7433,
6306,
3228,
7820,
2060,
1301,
1407,
5304,
8198,
5410,
6569,
7173,
1890,
1954,
9277,
8897,
7905,
3156,
7172,
8923,
2362,
6379,
7198,
2146,
7148,
9956,
3254,
#
#ノーベル賞銘柄対象(プラスそれに関連するバイオ、化学、文学、物理銘柄)
3386, #コスモバイオ
4528, #小野薬品
# 4098,　既に記載済
2370, #メディネット
4564, #オンコセラピーサイエンス
4587, #ペプチドリーム
# 4004, #昭和電光　既に記載済
4974, #タカラバイオ
# 4506, #大日本住友製薬　既に記述済
# 4523, #エーザイ　既に記載済
4592, #サンバイオ
4565, #そーせいグループ
4583, #カイオム・バイオサイエンス
4572, #カルナ・バイオレンス
4875, #メディシノバ
2191, #テラ
7776, #セルシード
4978, #リプロセル
# 4588, #オンコリスバイオファーマ　既に記載済
# 6752, #パナソニック　既に記載済
# 6762, #TDK　既に記載済
6981, #村田製作所
3891, #日本高度紙工業
# 6674, #ジーエスユアサ　既に記載済
4080, #田中化学研究所
4217, #日立化成
4027, #テイカ
4028, #石原産業
# 5801, #古川電気工業　既に記載済
# 5802, #住友電気工業　既に記載済
# 6503, #三菱電気 既に記載済
# 7013, #IHI　既に記載済
# 9022, #JR東海　既に記載済
9978, #文教堂グループホールディングス
3159, #丸善CHIホールディングス
4593, #ヘリオス
4563, #アンジェス
6955, #FDK

4078, #堺化学
5727, #邦チタ
3864, #三菱紙
4112, #保土谷
4205, #ゼオン
4215, #タキロンＣＩ
5998, #アドバネクス
6568, #神戸天然物化
6965, #ホトニクス
7915, #ＮＩＳＳＨＡ
7966, #リンテック
8101, #クレオス

#追加銘柄
7238,
6191,

#藤本先生セミナー
3277,
4245,
6062,
9263,
2485,
6625,
3054,
3021,
9889,
6195,

# 20199月高沢銘柄
3912,
3914,
4428,
4431,
3997,
6552,
3137,
3446,
3454,
# 3482,
3682,
3901,

# ミサイル関連銘柄
6072,
7721,
6946,
6203,
7963,
7980,
4274,
6208,

1557,  #SPDR S&P500 ETF 指数
7777,  #スリーディーマトリックス ノーベル賞銘柄
5287,  #イトーヨーギョー
6171,  #土木管理総合試験
9678,  #カナモト
5807,  #東京特殊電線
4825,  #ウェザーニューズ
1914,  #日本基礎技術
4318,  #クイック 人材銘柄
2410,  #キャリアデザインセンター 人材銘柄
2163,  #アルトナー 人材銘柄
6551,  #ツナググループHLDGS
4848,  #フルキャストホールディングス
6533,  #ORCHESTRAHLDGS
4557,  #医学生物学研究所
6502,  #東芝
7936,  #アシックス
8111,  #ゴールドウィン
7058,  #共栄セキュリティーサービス
3498,  #霞ヶ関キャピタル

#RFID
7862,  #トッパン・フォームズ     電子タグ(ＲＦＩＤ)
6145,  #日特エンジニアリング     電子タグ(ＲＦＩＤ)
5162,  #朝日ラバー             無溶剤接着の新型電子タグ(ＲＦＩＤ)
6287,  #サトーホールディングス   電子タグ(ＲＦＩＤ)スキャナ機器
6664,  #オプトエレクトロニクス   電子タグ(ＲＦＩＤ)スキャナ機器
3753,  #フライトホールディングス 電子タグ(ＲＦＩＤ)スキャナ機器
6945,  #富士通フロンテック      無人レジ(セルフレジ)メーカー
3837,  #アドソル日進           無人レジ　ハンズフリー認証システム
9972,  #アルテック             包装・印刷　電子タグ(ＲＦＩＤ)の埋込み包装
7919,  #野崎印刷紙業           包装・印刷　電子タグ(ＲＦＩＤ)の埋込み包装
6666,  #リバーエレテック        水晶デバイス(電子タグに必須の部品)

4998,  #フマキラー             害虫銘柄
1813,  #不動テトラ             災害銘柄
9699,  #西尾レントオール        災害銘柄
3030,  #ハブ






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



  test = CSV.open("#{mCode}.csv","a") do |csv|
    ##aデータ更新#
    ##w新規でテーブル作成

  csv << [
          mCode.to_s,
          # Date.today.strftime("%Y-%m-%d"),
          "2019-11-20",
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
          a13,
        ]

  end

  c = rand(1..5)

  sleep(c)

}
