
require 'slack-ruby-client'
require 'nokogiri'
require 'open-uri'


Slack.configure do |config|
  config.token = 'xoxb-594546582053-853798190228-TNHYQwyhJI3U0Su28q4MlD8e'
end

# #F&gindex取得
# url = "https://money.cnn.com/data/fear-and-greed/"
# charset = nil
#
# html = open(url) do |f|
#     charset = f.charset
#     f.read
# end
#
# doc = Nokogiri::HTML.parse(html, nil, charset)


# results = doc.xpath("//*[@id='needleChart']/ul").inner_text


client = Slack::Web::Client.new
client.chat_postMessage(channel: '#taketaketest', text: 'aaaa')
