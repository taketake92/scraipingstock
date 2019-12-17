
require 'slack-ruby-client'
require 'nokogiri'
require 'open-uri'


Slack.configure do |config|
  config.token = 'xoxb-594546582053-861117833618-xfVZo4yVEHMrQ7Q9dFpwIlbA'
end

#F&gindex取得
url = "https://money.cnn.com/data/fear-and-greed/"
charset = nil

html = open(url) do |f|
    charset = f.charset
    f.read
end

doc = Nokogiri::HTML.parse(html, nil, charset)

p doc

results1 = doc.xpath("//*[@id='needleChart']/ul/li[1]").inner_text
results2 = doc.xpath("//*[@id='needleChart']/ul/li[2]").inner_text
results3 = doc.xpath("//*[@id='needleChart']/ul/li[3]").inner_text
results4 = doc.xpath("//*[@id='needleChart']/ul/li[4]").inner_text
results5 = doc.xpath("//*[@id='needleChart']/ul/li[5]").inner_text


client = Slack::Web::Client.new
client.chat_postMessage(channel: '#taketaketest', text: results1 + "\n" +
                                                          results2 + "\n" +
                                                          results3 + "\n" +
                                                          results4 + "\n" +
                                                          results5)
