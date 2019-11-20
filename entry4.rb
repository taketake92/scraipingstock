require "google_drive"
require 'nokogiri'
require 'open-uri'
require 'date'
require 'CSV'

mCodes = [
'0000',
# 2002,
# 2269,
# 2282,
# 2501,
# 2502,
# 2503,
# 2531,
# 2801,
# 2802,
# 2871,
# 2914,
# 3101,
# 3103,
# 3401,
# 3402,
# 3861,
# 3863,
# 3405,
# 3407,
# 4004,
# 4005,
# 4021,
# 4042,
# 4043,
# 4061,
# 4063,
# 4183,
# 4188,
# 4208,
# 4272,
# 4452,
# 4631,
# 4901,
# 4911,
# 4151,
# 4502,
# 4503,
# 4506,
# 4507,
# 4519,
# 4523,
# 4568,
# 4578,
# 5019,
# 5020,
# 5101,
# 5108,
# 5201,
# 5202,
# 5214,
# 5232,
# 5233,
# 5301,
# 5332,
# 5333,
# 5401,
# 5406,
# 5411,
# 5541,
# 3436,
# 5703,
# 5706,
# 5707,
# 5711,
# 5713,
# 5714,
# 5801,
# 5802,
# 5803,
# 5901,
# 5631,
# 6103,
# 6113,
# 6301,
# 6302,
# 6305,
# 6326,
# 6361,
# 6367,
# 6471,
# 6472,
# 6473,
# 7004,
# 7011,
# 7013,
# 3105,
# 6479,
# 6501,
# 6503,
# 6504,
# 6506,
# 6645,
# 6674,
# 6701,
# 6702,
# 6703,
# 6724,
# 6752,
# 6758,
# 6762,
# 6770,
# 6841,
# 6857,
# 6902,
# 6952,
# 6954,
# 6971,
# 6976,
# 6988,
# 7735,
# 7751,
# 7752,
# 8035,
# 7003,
# 7012,
# 7201,
# 7202,
# 7203,
# 7205,
# 7211,
# 7261,
# 7267,
# 7269,
# 7270,
# 7272,
# 4543,
# 4902,
# 7731,
# 7733,
# 7762,
# 7832,
# 7911,
# 7912,
# 7951,
# 1332,
# 1333,
# 1605,
# 1721,
# 1801,
# 1802,
# 1803,
# 1808,
# 1812,
# 1925,
# 1928,
# 1963,
# 2768,
# 8001,
# 8002,
# 8015,
# 8031,
# 8053,
# 8058,
# 3086,
# 3099,
# 3382,
# 8028,
# 8233,
# 8252,
# 8267,
# 9983,
# 7186,
# 8303,
# 8304,
# 8306,
# 8308,
# 8309,
# 8316,
# 8331,
# 8354,
# 8355,
# 8411,
# 8601,
# 8604,
# 8628,
# 8630,
# 8725,
# 8729,
# 8750,
# 8766,
# 8795,
# 8253,
# 3289,
# 8801,
# 8802,
# 8804,
# 8830,
# 9001,
# 9005,
# 9007,
# 9008,
# 9009,
# 9020,
# 9021,
# 9022,
# 9062,
# 9064,
# 9101,
# 9104,
# 9107,
# 9202,
# 9301,
# 9412,
# 9432,
# 9433,
# 9437,
# 9613,
# 9984,
# 9501,
# 9502,
# 9503,
# 9531,
# 9532,
# 2432,
# 4324,
# 4689,
# 4704,
# 4751,
# 4755,
# 6098,
# 6178,
# 9602,
# 9681,
# 9735,
# 9766,
# 3990,
# 3556,
# 3665,
# 3987,
# 4320,
# 4427,
# 4726,
# 5704,
# 6184,
# 6754,
# 6963,
# 7974,
# 9467,
# 3110,
# 3038,
# 2326,
# 2384,
# 3046,
# 3064,
# 3349,
# 3479,
# 3687,
# 6920,
# 7309,
# 7550,
# 7564,
# 7816,
# 9934,
# 1357,
# 4588,
# 4344,
# 7748,
# 3906,
# 3356,
# 9434,
# 6541,
# 3092,
# 3967,
# 6861,
# 3627,
# 1570,
# 6033,
# 6071,
# 3652,
# 3962,
# 4238,
# 6155,
# 6246,
# #ここからスクリーニング結果
# 7827,
# 3477,
# 8903,
# 3467,
# 3604,
# 5697,
# 1795,
# 6580,
# 8885,
# 1844,
# 3695,
# 8886,
# 6335,
# 6694,
# 6867,
# 3241,
# 3524,
# 6267,
# 6570,
# 3750,
# 3814,
# 9625,
# 2974,
# 6757,
# 3461,
# 1739,
# 3294,
# 7781,
# 3647,
# 2436,
# 3489,
# 6032,
# 4356,
# 1431,
# 2599,
# 1826,
# 5941,
# 4952,
# 1994,
# 2162,
# 4323,
# 8893,
# 3440,
# 6362,
# 1420,
# 6164,
# 3538,
# 8144,
# 4098,
# 3550,
# 3323,
# 2153,
# 7559,
# 8732,
# 3934,
# 5922,
# 3919,
# 6159,
# 7722,
# 8887,
# 6403,
# 6059,
# 7059,
# 2676,
# 9233,
# 9057,
# 7539,
# 1960,
# 4987,
# 3452,
# 6331,
# 1726,
# 7745,
# 5290,
# 3482,
# 8929,
# 2301,
# 6535,
# 5013,
# 6089,
# 6677,
# 8917,
# 3772,
# 9270,
# 3546,
# 6257,
# 7199,
# 8877,
# 7433,
# 6306,
# 3228,
# 7820,
# 2060,
# 1301,
# 1407,
# 5304,
# 8198,
# 5410,
# 6569,
# 7173,
# 1890,
# 1954,
# 9277,
# 8897,
# 7905,
# 3156,
# 7172,
# 8923,
# 2362,
# 6379,
# 7198,
# 2146,
# 7148,
# 9956,
# 3254,
# #
# #ノーベル賞銘柄対象(プラスそれに関連するバイオ、化学、文学、物理銘柄)
# 3386, #コスモバイオ
# 4528, #小野薬品
# # 4098,　既に記載済
# 2370, #メディネット
# 4564, #オンコセラピーサイエンス
# 4587, #ペプチドリーム
# # 4004, #昭和電光　既に記載済
# 4974, #タカラバイオ
# # 4506, #大日本住友製薬　既に記述済
# # 4523, #エーザイ　既に記載済
# 4592, #サンバイオ
# 4565, #そーせいグループ
# 4583, #カイオム・バイオサイエンス
# 4572, #カルナ・バイオレンス
# 4875, #メディシノバ
# 2191, #テラ
# 7776, #セルシード
# 4978, #リプロセル
# # 4588, #オンコリスバイオファーマ　既に記載済
# # 6752, #パナソニック　既に記載済
# # 6762, #TDK　既に記載済
# 6981, #村田製作所
# 3891, #日本高度紙工業
# # 6674, #ジーエスユアサ　既に記載済
# 4080, #田中化学研究所
# 4217, #日立化成
# 4027, #テイカ
# 4028, #石原産業
# # 5801, #古川電気工業　既に記載済
# # 5802, #住友電気工業　既に記載済
# # 6503, #三菱電気 既に記載済
# # 7013, #IHI　既に記載済
# # 9022, #JR東海　既に記載済
# 9978, #文教堂グループホールディングス
# 3159, #丸善CHIホールディングス
# 4593, #ヘリオス
# 4563, #アンジェス
# 6955, #FDK
#
# 4078, #堺化学
# 5727, #邦チタ
# 3864, #三菱紙
# 4112, #保土谷
# 4205, #ゼオン
# 4215, #タキロンＣＩ
# 5998, #アドバネクス
# 6568, #神戸天然物化
# 6965, #ホトニクス
# 7915, #ＮＩＳＳＨＡ
# 7966, #リンテック
# 8101, #クレオス
#
# #追加銘柄
# 7238,
# 6191,
#
# #藤本先生セミナー
# 3277,
# 4245,
# 6062,
# 9263,
# 2485,
# 6625,
# 3054,
# 3021,
# 9889,
# 6195,
]

mCodes.each{|mCode|
  #テンプレートからコピーして新規にsheet作成
  #テンプレート読み込み
  templeate_sheets = ""
  copy_sheets = ""
  templeate_sheets = session.spreadsheet_by_key("1nQkxXvslr3nNODixgt2ETkNjB2ZyLbG5oWTiFJCauvk")
  #テンプレート作成
  p mCode
  copy_sheets = templeate_sheets.copy("#{mCode}")
  #テンプレートのid取得
  sheet_id = copy_sheets.id
  #テンプレートでシートインスタンス
  sheets = session.spreadsheet_by_key(sheet_id).worksheets[0]

  row = 2
  CSV.foreach("../#{mCode}.csv") do |line|

  	sheets[row,1] = line[0] #銘柄コード
  	sheets[row,2] = line[1] #日付
  	sheets[row,3] = line[10] #前日比
  	sheets[row,4] = line[11] #前日比(㌫)
  	sheets[row,5] = line[2] #始値
  	sheets[row,6] = line[3] #高値
  	sheets[row,7] = line[4] #高値(時刻)
  	sheets[row,8] = line[5] #安値
  	sheets[row,9] = line[6] #安値(時刻)
  	sheets[row,10] = line[7] #終値
  	sheets[row,11] = line[8] #出来高
  	sheets[row,12] = line[9] #売買代金
  	sheets[row,13] = line[13] #約定回数
  	sheets[row,15] = line[12] #VWAP
  	sheets[row,14] = line[13] #時価総額
  	row += 1

  end
  sheets.save
}