
require 'slack-ruby-client'

Slack.configure do |config|
  config.token = 'xoxb-594546582053-853798190228-TNHYQwyhJI3U0Su28q4MlD8e'
end

client = Slack::Web::Client.new
client.chat_postMessage(channel: '#taketaketest', text: 'テストメール')
