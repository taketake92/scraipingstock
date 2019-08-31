require "google_drive"
require 'nokogiri'
require 'open-uri'
require 'date'
require 'CSV'


#html分解
def nokogiri(url)
  charset = nil
  html = open(url) do |f|
      charset = f.charset
      f.read
  end
  p url
  return Nokogiri::HTML.parse(html, nil, charset)

end

def rating(doc)
  result_hash = {}

  # cbs1 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[2]/th").inner_text
  # cbs2 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[3]/th").inner_text
  # cbs3 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[4]/th").inner_text
  # cbs4 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[5]/th").inner_text
  # cbs5 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[6]/th").inner_text

  # result = {"cbs": cbs1,"cbs": cbs2, "cbs": cbs3, "cbs": cbs4, "cbs": cbs5}

  cbs_hash = {cbs: {cbs1: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[2]/th").inner_text,
                cbs2: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[3]/th").inner_text,
                cbs3: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[4]/th").inner_text,
                cbs4: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[5]/th").inner_text,
                cbs5: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[6]/th").inner_text
                }}
 # p cbs_hash


  # p result

  # p cbs1
  # p cbs2
  # p cbs3
  # p cbs4
  # p cbs5

  #売上高
  aos1 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[2]/td[1]").inner_text
  aos1 = aos1.include?(",") ? aos1.delete!(",") : aos1
  aos2 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[3]/td[1]").inner_text
  aos2 = aos2.include?(",") ? aos2.delete!(",") : aos2
  aos3 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[4]/td[1]").inner_text
  aos3 = aos3.include?(",") ? aos3.delete!(",") : aos3
  aos4 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[5]/td[1]").inner_text
  aos4 = aos4.include?(",") ? aos4.delete!(",") : aos4
  aos5 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[6]/td[1]").inner_text
  aos5 = aos5.include?(",") ? aos5.delete!(",") : aos5

  aos_hash = {aos: {aos1: aos1,
                aos2: aos2,
                aos3: aos3,
                aos4: aos4,
                aos5: aos5
                }}

  p aos_hash

  # result = {"aos1": aos1, "aos2": aos2, "aos3": aos3, "aos4": aos4, "aos5": aos5}

  #営業益
  oi1 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[2]/td[2]").inner_text
  oi1 = oi1.include?(",") ? oi1.delete!(",") : oi1
  oi2 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[3]/td[2]").inner_text
  oi2 = oi2.include?(",") ? oi2.delete!(",") : oi2
  oi3 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[4]/td[2]").inner_text
  oi3 = oi3.include?(",") ? oi3.delete!(",") : oi3
  oi4 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[5]/td[2]").inner_text
  oi4 = oi4.include?(",") ? oi4.delete!(",") : oi4
  oi5 = doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[6]/td[2]").inner_text
  oi5 = oi5.include?(",") ? oi5.delete!(",") : oi5

  p oi1
  p oi2
  p oi3
  p oi4
  p oi5

  oi_hash = {oi: {oi1: oi1,
                oi2: oi2,
                oi3: oi3,
                oi4: oi4,
                oi5: oi5
                }}

  #売上高フラグ
  aosFlg = (aos1.to_i < aos2.to_i) && (aos2.to_i < aos3.to_i) && (aos3.to_i < aos4.to_i) && (aos4.to_i < aos5.to_i) ? true : false

  #営業履歴フラグ
  oiFlg = (oi1.to_i < oi2.to_i) && (oi2.to_i < oi3.to_i) && (oi3.to_i < oi4.to_i) && (oi4.to_i < oi5.to_i) ? true : false


  #長期的増加
  oidFlg =  ((2 * oi1.to_i) < oi5.to_i) ? true : false

  #成長性レーティング
  if ((aosFlg = true && oiFlg == true) && oidFlg == true) then
    gpr = 7
  elsif((aosFlg = true && oiFlg == true) && oidFlg == false) then
    gpr = 5
  elsif((aosFlg = false || oiFlg == false) && oidFlg == true) then
    gpr = 4
  elsif((aosFlg = false || oiFlg == false) && oidFlg == false) then
    gpr = 2
  elsif((aosFlg = false && oiFlg == false) && oidFlg == false) then
    gpr = -3
  end

  gpr_hash = {gpr: gpr}

  p gpr

  #自己資本比率
  car = doc.xpath("//*[@id='finance_box']/table[4]/tbody/tr[4]/td[2]").inner_text.to_i

  if car >= 60 then
    carr = 3
  elsif car >= 40 && car < 60 then
    carr = 2
  elsif car >= 20 && car < 40 then
    carr = 0
  else
    carr = -1
  end

  p = carr

  carr_hash = {carr: carr}

  #ROE
  roe = doc.xpath("//*[@id='finance_box']/div[9]/table/tbody/tr[3]/td[4]").inner_text.to_i

  if roe > 15 then
    roer = 3
  elsif roe >= 8 && roe < 15 then
    roer = 2
  elsif roe >= 5 && roe < 8 then
    roer = 1
  else
    roer = -1
  end

  p roer

  roer_hash = {roer: roer}

  return result_hash.merge(cbs_hash)
                            .merge(aos_hash)
                            .merge(oi_hash)
                            .merge(gpr_hash)
                            .merge(carr_hash)
                            .merge(roer_hash)

end

def rating2(doc)

  per = doc.xpath("//*[@id='JSID_stockInfo']/div[1]/div[1]/div[1]/div[2]/ul/li[2]/span[2]").inner_text.delete!(" 倍")
  per = per.to_f

  p per

  if per <= 8 then
    perr = 7
  elsif per <= 12 && per > 8 then
    perr = 5
  elsif per <= 8 && per > 15 then
    perr = 3
  else
    perr = 2
  end

  p perr

end

def rating3(doc)
  if doc.to_s.include?('ico_w_01_s.v2.png') then
    rar = 3
  elsif doc.to_s.include?('ico_w_02_s.v2.png') then
    rar = 2
  elsif doc.to_s.include?('ico_w_03_s.v2.png') then
    rar = 0
  else
    rar = "null"
  end
  p rar
end

session = GoogleDrive::Session.from_config("config.json")

mCodes = [
  1420,
1431,
# 1726,
# 1739,
# 1795,
# 1826,
# 1844,
# 1960,
# 1994,
# 2153,
# 2162,
# 2301,
# 2436,
# 2599,
# 2676,
# 2974,
# 3241,
# 3294,
# 3323,
# 3440,
# 3440,
# 3452,
# 3461,
# 3467,
# 3477,
# 3482,
# 3489,
# 3524,
# 3538,
# 3546,
# 3550,
# 3604,
# 3647,
# 3695,
# 3750,
# 3772,
# 3814,
# 3919,
# 3934,
# 4098,
# 4323,
# 4356,
# 4952,
# 4987,
# 5013,
# 5290,
# 5697,
# 5922,
# 5941,
# 6032,
# 6059,
# 6089,
# 6159,
# 6164,
# 6257,
# 6267,
# 6306,
# 6331,
# 6335,
# 6362,
# 6403,
# 6535,
# 6570,
# 6580,
# 6677,
# 6694,
# 6757,
# 6867,
# 7059,
# 7199,
# 7433,
# 7539,
# 7559,
# 7722,
# 7745,
# 7781,
# 7827,
# 8732,
# 8877,
# 8885,
# 8886,
# 8887,
# 8893,
# 8903,
# 8917,
# 8929,
# 9057,
# 9233,
# 9625,
]

a = 0




mCodes.each{|mCode|

# https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxx/
# 事前に書き込みたいスプレッドシートを作成し、上記スプレッドシートのURL(「xxx」の部分)を以下のように指定する
sheets = session.spreadsheet_by_key("1Sw1AlOXEZ6HCnnN-NnPduIrDpi7u3zOaLf7k_wCyaDo").worksheets[1]
# sheets = session.spreadsheet_by_key("1IUyY4mHmAwSBkNJks9WbpIAlGKI4foQM9BfyNgAzl4c").worksheets[0]



startRow = 3 #行

  url = "https://kabutan.jp/stock/finance?code=#{mCode}"

  #url分解
  doc = nokogiri(url)

  #rating結果
  rating1 = rating(doc)

  url = "https://www.nikkei.com/nkd/company/?scode=#{mCode}"

  #url分解
  doc = nokogiri(url)

  #予想PER
  rating2 = rating2(doc)

  #直近の業績予報
  url = "https://kabuyoho.ifis.co.jp/index.php?action=tp1&sa=report&bcode=#{mCode}"

  #url分解
  doc = nokogiri(url)

  #予想PER
  rating3 = rating3(doc)



  # スプレッドシートへの書き込み
  # sheets[startRow,1] = Date.today.strftime("%Y/%m/%d") #日付
  sheets[startRow,2] = "aaaa"
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



  # シートの保存
  sheets.save



#   test = CSV.open("#{mCode}.csv","a") do |csv|
#     ##翌日wをaに変えて上書き#
#
#   csv << [
#           mCode.to_s,
#           # Date.today.strftime("%Y-%m-%d"),
#           "2019-08-23",
#           a1,
#           a2,
#           a3,
#           a4,
#           a5,
#           a6,
#           a7,
#           a8,
#           a9,
#           a10,
#           a11,
#           a12,
#         ]
#
# end
#
# c = rand(30..50)
#
# sleep(c)

}
