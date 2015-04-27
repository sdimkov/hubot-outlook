phantom = require('phantom')

{Adapter, TextMessage, User} = require 'hubot'


# Configuration
reconnectInterval = false
if process.env.HUBOT_OUTLOOK_RECONNECT
  reconnectInterval = parseInt process.env.HUBOT_OUTLOOK_RECONNECT
  if reconnectInterval < 10
    console.error "HUBOT_OUTLOOK_RECONNECT is the adapter reconnect interval in minutes!"
    console.error "Minimum reconnect interval is 10 minutes!"
    process.exit 1


class Outlook extends Adapter


  run: ->
    self = @
    @connect ->
      self.emit 'connected'
      self.robot.logger.info 'Outlook adapter connected!'
    if reconnectInterval
      setInterval ( -> self.reconnect()), reconnectInterval * 60 * 1000
      @robot.logger.info "Outlook adapter configured to reconnect every #{reconnectInterval} minutes"


  reconnect: ->
    self = @
    @oldPhantom = @phantom
    @connect ->
      self.oldPhantom.exit 0
      delete self.oldPhantom
      self.emit 'reconnected'
      self.robot.logger.info 'Outlook adapter re-connected!'


  close: () ->
    @phantom.exit 0
    delete @phantom
    try
      if @oldPhantom
        @oldPhantom.exit 0
        delete @oldPhantom
    catch err
      @robot.logger.error "Unable to kill old phantom (during reconnecting): #{err}\n#{err.stack}"


  reply: (envelope, strings...) ->
    # Only prefix replies in group chats
    if envelope.user.room.indexOf('19:') is 0
      @robot.logger.debug 'reply: adding receiver prefix to all strings'
      strings = strings.map (s) -> "#{envelope.user.nickname || envelope.user.name}: #{s}"
    else
      @robot.logger.debug 'reply: replying in personal chat ' + envelope.user.room
    @send envelope, strings...


  connect: (callback) ->
    self = @
    phantom.create (ph) ->
      self.phantom = ph
      ph.createPage (page) ->
        page.set 'settings.userAgent',
          'Mozilla/5.0 (Windows NT 6.1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/41.0.2228.0 Safari/537.36'
        page.open 'http://live.com', (status) ->
          self.robot.logger.debug 'phantomjs opened live.com with status', status

          # Login to Outlook chat
          setTimeout (->
            page.evaluate ((username, password) ->
              document.querySelector('input[name=login]').value = username
              document.querySelector('input[name=passwd]').value = password
              document.querySelector('input[name=SI]').click()
            ), (->), process.env.HUBOT_OUTLOOK_USERNAME, process.env.HUBOT_OUTLOOK_PASSWORD
          ), 500
          
          setTimeout (->
            page.evaluate ->

              # Intercept each message upon arrival and forward to phantomjs
              @msgReceivedHandler = $WLXIM.MLM.Conversation::onMessageReceived
              $WLXIM.MLM.Conversation::onMessageReceived = (msg) ->
                window.callPhantom
                  id: msg._message$1._id
                  text: msg._message$1._text$1
                  # timestamp: msg._message$1._timestamp
                  user: 
                    id: msg._message$1._sender._addressInfo.id
                    room: msg._message$1._conversationId
                msgReceivedHandler.apply this, arguments

              # Define a helper function for sending messages
              @sendMessage = (user, msg) ->
                format = new (MC.TextMessageFormatInfo)(0, 'Segoe UI', 0, false)
                randomId = Math.random().toString() + Math.random().toString()
                randomId = randomId.replace(/\.|0/g, '').substring(0, 17)
                messageInfo = new (MC.MessageInfo)
                messageInfo.info = new (MC.TextMessageInfo)(msg, format)
                messageInfo.type = 1
                messageInfo.sender = $WLXIM.user.get_address()._addressInfo
                messageInfo.clientId = randomId
                $WLXIM.user._conversationSendMessage user, messageInfo

            # Catch forwarded messages from the page context
            page.set 'onCallback', (msg) ->
              user = self.robot.brain.userForId msg.user.id
              user.room = msg.user.room
              # Let robot know messages in personal chats are directed at him
              if user.room.indexOf('19:') isnt 0
                unless user.shell? and user.shell[user.room]
                  self.robot.logger.debug 'prefix personal message'
                  msg.text = self.robot.name + ': ' + msg.text
              # Provide the messages to the robot
              self.receive new TextMessage user, msg.text, msg.id

            # Sends messages back to outlook
            self.send = (envelope, strings...) ->
              page.evaluate ((room, strings) ->
                sendMessage room, msg for msg in strings
              ), (->), envelope.room, strings

            callback()

          ), 10000


exports.use = (robot) ->
  new Outlook robot

