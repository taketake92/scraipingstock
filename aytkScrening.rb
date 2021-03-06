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

  #概要
  # ov_hash = {ov: doc.xpath("//*[@id='kobetsu_right']/div[4]/table/tbody/tr[3]/td").inner_text}

  cbs_hash = {cbs: {cbs1: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[2]/th").inner_text,
                cbs2: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[3]/th").inner_text,
                cbs3: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[4]/th").inner_text,
                cbs4: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[5]/th").inner_text,
                cbs5: doc.xpath("//*[@id='finance_box']/div[6]/table/tbody/tr[6]/th").inner_text
                }}

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

  oi_hash = {oi: {oi1: oi1,
                oi2: oi2,
                oi3: oi3,
                oi4: oi4,
                oi5: oi5
                }}

  #売上高フラグ
  aosFlg = (aos1.to_i < aos2.to_i) == true &&
              (aos2.to_i < aos3.to_i) == true &&
              (aos3.to_i < aos4.to_i) == true &&
              (aos4.to_i < aos5.to_i) ? true : false

  #営業履歴フラグ
  oiFlg = (oi1.to_i < oi2.to_i) == true &&
              (oi2.to_i < oi3.to_i) == true &&
              (oi3.to_i < oi4.to_i) == true &&
              (oi4.to_i < oi5.to_i) == true ? true : false

  #長期的増加
  oidFlg =  ((2 * oi1.to_i) < oi5.to_i) ? true : false

  moji1 = /^[-]?[0-9]+$/ =~ aos1.to_s ? true : false
  moji2 = /^[-]?[0-9]+$/ =~ aos2.to_s ? true : false
  moji3 = /^[-]?[0-9]+$/ =~ aos3.to_s ? true : false
  moji4 = /^[-]?[0-9]+$/ =~ aos4.to_s ? true : false
  moji5 = /^[-]?[0-9]+$/ =~ aos5.to_s ? true : false
  moji6 = /^[-]?[0-9]+$/ =~ oi1.to_s ? true : false
  moji7 = /^[-]?[0-9]+$/ =~ oi2.to_s ? true : false
  moji8 = /^[-]?[0-9]+$/ =~ oi3.to_s ? true : false
  moji9 = /^[-]?[0-9]+$/ =~ oi4.to_s ? true : false
  moji10 = /^[-]?[0-9]+$/ =~ oi5.to_s ? true : false

  if(moji1 === false ||
    moji2 === false ||
    moji3 === false ||
    moji4 === false ||
    moji5 === false) ||
  (moji6 === false ||
    moji7 === false ||
    moji8 === false ||
    moji9 === false ||
    moji10 === false) then
    gpr = -99
  else
    if ((aosFlg == true && oiFlg == true) && oidFlg == true) then
      gpr = 7
    elsif((aosFlg == false || oiFlg == false) && oidFlg == true) then
      gpr = 5
    elsif((aosFlg == true && oiFlg == true) && oidFlg == false) then
      gpr = 4
    elsif((aosFlg == false || oiFlg == false) && oidFlg == false) then
      gpr = 2
    elsif((aosFlg == false && oiFlg == false) && oidFlg == false) then
      gpr = -3
    end
  end

  gpr_hash = {gpr: gpr}

  #自己資本比率
  car = doc.xpath("//*[@id='finance_box']/table[4]/tbody/tr[4]/td[2]").inner_text.to_f

  if car >= 60 then
    carr = 3
  elsif car >= 40 && car < 60 then
    carr = 2
  elsif car >= 20 && car < 40 then
    carr = 0
  else
    carr = -1
  end

  carr_hash = {car: "#{car}%" ,carr: carr}

  #ROE
  roe = doc.xpath("//*[@id='finance_box']/div[9]/table/tbody/tr[3]/td[4]").inner_text.to_f

  if roe > 15 then
    roer = 3
  elsif roe >= 8 && roe < 15 then
    roer = 2
  elsif roe >= 5 && roe < 8 then
    roer = 1
  else
    roer = -1
  end

  roer_hash = {roe: "#{roe}%", roer: roer}

  p result_hash.merge(cbs_hash)
                          .merge(aos_hash)
                          .merge(oi_hash)
                          .merge(gpr_hash)
                          .merge(carr_hash)
                          .merge(roer_hash)

  return result_hash.merge(cbs_hash)
                            .merge(aos_hash)
                            .merge(oi_hash)
                            .merge(gpr_hash)
                            .merge(carr_hash)
                            .merge(roer_hash)
end

#割安性取得
def rating2(doc)

  result_hash = {}

  per = doc.xpath("//*[@id='JSID_stockInfo']/div[1]/div[1]/div[1]/div[2]/ul/li[2]/span[2]").inner_text.delete!(" 倍")
  per = per.to_f

  if per <= 8 then
    perr = 7
  elsif per <= 12 && per > 8 then
    perr = 5
  elsif per <= 8 && per > 15 then
    perr = 3
  else
    perr = 2
  end

  return {per:"#{per}倍", perr:perr}

end

#直近の業績取得
def rating3(doc)
  result_hash = {}

  if doc.to_s.include?('ico_w_01_s.v2.png') then
    rar = "晴"
    rarr = 3
  elsif doc.to_s.include?('ico_w_02_s.v2.png') then
    rar = "曇"
    rarr = 2
  elsif doc.to_s.include?('ico_w_03_s.v2.png') then
    rar = "雨"
    rarr = 0
  else
    rar = "なし"
    rarr = 0
  end

  return {rar: rar, rarr: rarr}

end

session = GoogleDrive::Session.from_config("config.json")

mCodes = [
  1401,
  1431,
  1726,
  1736,
  1739,
  1758,
  1789,
  1826,
  1867,
  2060,
  2130,
  2151,
  2153,
  2162,
  2176,
  2301,
  2352,
  2404,
  2424,
  2429,
  2436,
  2480,
  2488,
  2676,
  2737,
  2777,
  2926,
  2981,
  3030,
  3031,
  3071,
  3073,
  3137,
  3150,
  3180,
  3189,
  3228,
  3245,
  3266,
  3271,
  3294,
  3323,
  3416,
  3418,
  3452,
  3454,
  3458,
  3463,
  3467,
  3475,
  3477,
  3482,
  3484,
  3486,
  3489,
  3491,
  3496,
  3498,
  3524,
  3538,
  3546,
  3547,
  3604,
  3625,
  3627,
  3656,
  3658,
  3665,
  3674,
  3676,
  3683,
  3695,
  3696,
  3698,
  3750,
  3772,
  3802,
  3810,
  3843,
  3849,
  3910,
  3919,
  3920,
  3922,
  3924,
  3934,
  3969,
  3981,
  3988,
  3997,
  4235,
  4245,
  4380,
  4390,
  4392,
  4398,
  4427,
  4433,
  4436,
  4437,
  4447,
  4481,
  4486,
  4491,
  4557,
  4699,
  4763,
  4770,
  4828,
  4832,
  4845,
  4952,
  4960,
  5189,
  5290,
  5922,
  6033,
  6059,
  6062,
  6070,
  6071,
  6077,
  6088,
  6089,
  6175,
  6191,
  6239,
  6255,
  6266,
  6306,
  6331,
  6362,
  6387,
  6391,
  6400,
  6403,
  6533,
  6539,
  6540,
  6557,
  6558,
  6560,
  6565,
  6570,
  6625,
  6629,
  6668,
  6734,
  6867,
  7035,
  7038,
  7044,
  7047,
  7048,
  7058,
  7059,
  7060,
  7073,
  7079,
  7080,
  7083,
  7090,
  7172,
  7191,
  7199,
  7213,
  7425,
  7433,
  7590,
  7674,
  7676,
  7685,
  7722,
  7743,
  7748,
  7781,
  7820,
  8023,
  8144,
  8704,
  8732,
  8877,
  8886,
  8887,
  8893,
  8918,
  8927,
  8929,
  9057,
  9270,
  9279,
  9381,
  9467,
  9625,
  9629,
  9742,
  9816,
  9820,
  9889,

]

count = 0

startRow = 3 #行

mCodes.each{|mCode|

  # https://docs.google.com/spreadsheets/d/xxxxxxxxxxxxxxxxxxxxxxxxxxx/
  # 事前に書き込みたいスプレッドシートを作成し、上記スプレッドシートのURL(「xxx」の部分)を以下のように指定する
  sheets = session.spreadsheet_by_key("1KBbdzO6TnZ2w1m5f1xN-YXNolliR5pUx1hWmAE7bBzg").worksheets[1]
  # sheets = session.spreadsheet_by_key("1IUyY4mHmAwSBkNJks9WbpIAlGKI4foQM9BfyNgAzl4c").worksheets[0]
  
  p sheets

  p mCode
  startRow = 3 #行
  startColumn = 3 #列
  # startColumn = startColumn + (5 * count)
  startColumn = startColumn + (5 * count)

  p startRow

  url = "https://kabutan.jp/stock/finance?code=#{mCode}"

  #url分解
  doc = nokogiri(url)

  #rating結果
  rating1 = rating(doc)
  p rating1

  url = "https://www.nikkei.com/nkd/company/?scode=#{mCode}"

  #url分解
  doc = nokogiri(url)

  #予想PER
  rating2 = rating2(doc)
  p rating2

  #直近の業績予報
  url = "https://kabuyoho.ifis.co.jp/index.php?action=tp1&sa=report&bcode=#{mCode}"

  #url分解
  doc = nokogiri(url)

  #予想PER
  rating3 = rating3(doc)
  p rating3


  # スプレッドシートへの書き込み
  # 概要
  # sheets[3,startRow - 1] = rating1[:ov]

  # 銘柄コード
  sheets[startColumn, startRow - 2] = mCode

  # 成長性/業績
  sheets[startColumn, startRow] = rating1[:cbs][:cbs1]
  sheets[startColumn + 1, startRow] = rating1[:cbs][:cbs2]
  sheets[startColumn + 2, startRow] = rating1[:cbs][:cbs3]
  sheets[startColumn + 3, startRow] = rating1[:cbs][:cbs4]
  sheets[startColumn + 4, startRow] = rating1[:cbs][:cbs5]

  p startRow

  # 成長性/売上高
  sheets[startColumn, startRow + 1] = rating1[:aos][:aos1]
  sheets[startColumn + 1, startRow + 1] = rating1[:aos][:aos2]
  sheets[startColumn + 2, startRow + 1] = rating1[:aos][:aos3]
  sheets[startColumn + 3, startRow + 1] = rating1[:aos][:aos4]
  sheets[startColumn + 4, startRow + 1] = rating1[:aos][:aos5]

  # 成長性/営業利益
  sheets[startColumn, startRow + 2] = rating1[:oi][:oi1]
  sheets[startColumn + 1, startRow + 2] = rating1[:oi][:oi2]
  sheets[startColumn + 2, startRow + 2] = rating1[:oi][:oi3]
  sheets[startColumn + 3, startRow + 2] = rating1[:oi][:oi4]
  sheets[startColumn + 4, startRow + 2] = rating1[:oi][:oi5]

  # 成長性/点数結果
  sheets[startColumn, startRow + 3] = rating1[:gpr]

  # 安定性/点数結果
  sheets[startColumn + 1, startRow + 4] = rating1[:car]
  sheets[startColumn + 3, startRow + 4] = rating1[:carr]

  # 効率性/効率性点数結果
  sheets[startColumn + 1, startRow + 5] = rating1[:roe]
  sheets[startColumn + 3, startRow + 5] = rating1[:roer]

  # 割安性/割安性点数結果
  sheets[startColumn + 1, startRow + 6] = rating2[:per]
  sheets[startColumn + 3, startRow + 6] = rating2[:perr]

  # 直近の業績/直近の業績点数結果
  sheets[startColumn + 1, startRow + 7] = rating3[:rar]
  # sheets[startColumn + 3, startRow + 7] = rating3[:rarr]

  # レーティング結果
  sheets[startColumn, startRow + 8] = rating1[:gpr].to_i +
                            rating1[:carr].to_i +
                            rating1[:roer].to_i +
                            rating2[:perr].to_i +
                            rating3[:rarr].to_i
  # # シートの保存
  sheets.save
  count += 1
}
