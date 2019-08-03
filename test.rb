require 'nokogiri'
require 'open-uri'
require "csv"

url = 'https://www.nikkei.com/nkd/company/?scode=3990'

charset = nil

html = open(url) do |f|
    charset = f.charset
    f.read
end

doc = Nokogiri::HTML.parse(html, nil, charset)

test = CSV.open('test.csv','w') do |csv|

  a1 = doc.xpath('//*[@id="JSID_stockInfo"]/div[1]/div[1]/div[1]/div[1]/ul/li[1]/span[2]').inner_text
  #   a1 = node1.inner_text
  # end
  #
  a2 = doc.xpath('//*[@id="JSID_stockInfo"]/div[1]/div[1]/div[1]/div[1]/ul/li[2]/span[2]').inner_text

  #
  # doc.xpath('//*[@id="JSID_stockInfo"]/div[1]/div[1]/div[1]/div[1]/ul/li[2]/span[1]').each do |node3|
  #   a3 = node3.inner_text
  # end
  #
  # doc.xpath('//*[@id="JSID_stockInfo"]/div[1]/div[1]/div[1]/div[1]/ul/li[3]/span[2]').each do |node4|
  #   p node4.inner_text
  # end
  #
  # doc.xpath('//*[@id="JSID_stockInfo"]/div[1]/div[1]/div[1]/div[1]/ul/li[3]/span[1]').each do |node5|
  #   p node5.inner_text
  # end
  #
  # doc.xpath('//*[@id="CONTENTS_MAIN"]/div[3]/dl[1]/dd').each do |node6|
  #   p node6.inner_text
  # end
  #
  # doc.xpath('//*[@id="CONTENTS_MAIN"]/div[3]/dl[2]/dd/span[1]/span').each do |node7|
  #   p node7.inner_text
  # end
  #
  # doc.xpath('//*[@id="JSID_stockInfo"]/div[1]/div[1]/div[1]/div[2]/ul/li[1]/span[2]').each do |node8|
  #   p node8.inner_text
  # end
  #
  csv << ["3990",a1,a2]

end
