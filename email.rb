require "bundler"
Bundler.require(:default)
require 'net/imap'
require 'net/pop'
require_relative 'tellurium_driver.rb' 


# class to scrape a given email address for messages that match a given subject, sender,
# and body. Supports most major IMAP providers (gmail, att, yahoo, aol, ymail). Every 
# other provider is supported by pop3, which can only access the inbox, not the spam.

# USAGE
# 1. Verify that your provider's information is in the @@info hash
# 2. Create an new instance of the Email class with your username and password
#         email = Email.new(*****@***.com,<password>)
# 3. Run the search method on Email. This supports :subject, :from, :body
#    and will return an array of size two with the first element as the inbox
#    matching message count and the second as the spambox matching message count
#         results = email.search({subject:"foo",from:"foo@foo.com",body: "foo",number: 10,seen:"SEEN",date:"1-Apr-2003"})
#  note, options for number 0, infiniti. options for seen: "SEEN" or "UNSEEN" (read/unread basically), date must be in format above

class Email
include Enumerable
@@info = {
    gmail: {
      folders: ['INBOX','[Gmail]/Spam'], 
      method: "imap",
      url: "imap.gmail.com",
      ssl: true,
      port: 993
    },
    att: {
      folders: ['INBOX',"Bulk Mail"],
      method: "imap",
      url: "imap.mail.yahoo.com",
      ssl: true,
      port: 993
      },
    yahoo: {
      folders: ['INBOX',"Bulk Mail"],
      method: "imap",
      url: "imap.mail.yahoo.com",
      ssl: true,
      port: 993
    },
    ymail: {
      folders: ['INBOX',"Bulk Mail"],
      method: "imap",
      url: "imap.mail.yahoo.com",
      ssl: true,
      port: 993
    },
    aol: {
      folders: ['INBOX','Spam'],
      method: "imap",
      url: "imap.aol.com",
      ssl: false,
      port: 143
      },
    live: {
      method: "pop3",
      url: "pop3.live.com"
    },
    msn:{
      method: "pop3",
      url: "pop3.live.com"
    },
    hotmail: {
      method: "pop3",
      url: "pop3.live.com"
    },
    verizon: {
      method: "pop3",
      url: "incoming.verizon.net"
      },
    comcast: {
      method: "pop3",
      url: "mail.comcast.net"
     }
}

  def initialize(*args)
    @username,@password = args
    @provider = @username.split('@')[1].split('.')[0].to_sym #basically just stripping the provider name, like user@yahoo.com returns :yahoo
    @provider_info = @@info[@provider] #get the login information for the givin provider
    if(@provider_info[:method] == "imap")
      @imap = Net::IMAP.new(@provider_info[:url],@provider_info[:port],@provider_info[:ssl])
      @imap.login(@username, @password)
      
    elsif(@provider_info[:method] == "pop3") 
      @pop = Net::POP3.new(@provider_info[:url])
      @pop.enable_ssl(OpenSSL::SSL::VERIFY_NONE)  
           
    else
      raise "Invalid Method, please use imap or pop3"
    end
 
  end

  #incredibly long winded way of extracting "From" address. Takes a pop message 'm'
  def self.parse_pop_from(m)
    head = m.header.split(%r{\n}) #split the header into an array of strings
    return head.grep(/From/)[0].split('<')[1].split('>') #Search for the part of the array with the from in it and strip the email out from between these things: <>
  end
  #takes a pop message 'm' and returns the subject of the message
  def self.parse_pop_subject(m)
    head = m.header.split(%r{\n}) #split the header into an array of strings
    return head.grep(/Subject/)[0].split(':') #Search for the part of the header array with subject. Exctract the subject
  end

  #Similar to above, pop the message in html, search for where it says the date, take all of the html after that
  #Search that html for what you want in the body. Note that everything after the dateline excludes everything but
  #the body of the message. Takes a pop message 'm'
  def self.parse_pop_body(m)
    html_msg = m.pop.split(%r{\n})
    dateline = html_msg.grep(/Date/)
    date_index = html_msg.find_index(dateline[0])

    return html_msg[date_index..(html_msg.size-1)]
  end

  def search_pop3(subject,from,body,number)
      count = []
      count.push 0
      @pop.start(@username,@password)
      @pop.each_mail do |m|
        if number != 0
          if Email.parse_pop_from(m).grep(/#{from}/)[0] !=nil and Email.parse_pop_subject(m).grep(/#{subject}/)[0] !=nil and Email.parse_pop_body(m).grep(/#{body}/)[0] !=nil
            count[0] += 1 #since we are using pop, it will still return an array so it's similar to the imap function, but the array only has a size of one. This is just for convenience sake, it can be changed
          end 
         
          number -= 1
        end

      end
      return count 
  end

  #opens the specified folder in the imap client
  def get_from_folder(folder)
    @imap.examine(folder)
    @imap.check
  end

  def search_imap(subject,from,body,number,date,seen)
    #make each attribute an empty string if it's nil
    body = "" unless body
    from = "" unless from
    subject = "" unless subject
          results = [0,0]
    folders = @provider_info[:folders]
    i = 0

    folders.each do |folder|
      self.get_from_folder(folder)
      
      #search for the given subject,from, and body and store in the results array. 
      #results[0] is always inbox, results[1] is spam      
      #although it looks sloppy, a separate if statement for each part of the logic ensures that 
      #time isn't wasted grabbing from lines or bodys if the subject doesn't match
      #,["SINCE","1-Apr-2003"]
      search = ["SINCE",date,seen] if seen != ""
      search = ["SINCE",date] if seen == ""
      @imap.search(search).each do |message_id|

        if number != 0
          message_subject = @imap.fetch(message_id, "BODY[HEADER.FIELDS (SUBJECT)]")[0].to_s
          
          if message_subject.include?(subject)
            message_from = @imap.fetch(message_id, "BODY[HEADER.FIELDS (FROM)]")[0].to_s
            
            if message_from.include?(from)
              message_body = @imap.fetch(message_id,'BODY[TEXT]')[0].attr['BODY[TEXT]']
              
              results[i] +=1 if message_body.include?(body)
            end

          end
          number -=1
        end
        
      end

      i+=1
    end 

    return results
  end  
 
  def search_owa(subject,from,body,number)
    return nil if @provider == :comcast
    browser = TelluriumDriver.new("1","2","3") 
    begin
      browser.driver.get("http://mail.live.com")
      browser.driver.find_element(:id,"i0118").send_keys(@password)
      browser.driver.find_element(:css,"input[type=email]").send_keys(@username)
      browser.driver.find_element(:id,"idSIButton9").click
      sleep(10)
      junk_element = browser.driver.find_element(:css,"#folderList [nm=Junk]")		
      junk_element.click
      sleep(4) #this thing is seriously finicky, not even the best wait_to_appear function helps. You have to wait
      browser.driver.find_element(:css,"#contentRight #messageListContentContainer .ia_hc").click
      browser.wait_to_appear(:id,"mpf0_MsgContainer")			
      
      @number_found = 0
      hit_the_one_after = 0
      #hit_the_one_after is used to count one after the thing isn't displayed, that way we hit the last email
      while(browser.driver.find_element(:css,".i_rm_n").displayed? or hit_the_one_after == 1) and number!=0  do
        body = browser.driver.find_element(:id,"mpf0_MsgContainer").text
        
        unsafe_email_elements = browser.driver.find_elements(:css,"span.UnsafeSenderEmail")

        if( unsafe_email_elements.size > 0 ) #if it's an unknown sender, the email will be directly available
          from = unsafe_email_elements[0].text

        else #it's a known contact, click their name and get their email
          browser.driver.find_element(:id,"rmic1_senderName").click
          from = browser.driver.find_element(:css,"div.c_ic_menu_sub").text
        end 

        subject = browser.driver.find_element(:css,".ReadMsgSubject").text
        @number_found += 1 if body.include?("woohoo") and subject.include?("woohoo")

        browser.driver.find_element(:css,".i_rm_n").click unless hit_the_one_after == 1

        hit_the_one_after += 1 unless browser.driver.find_element(:css,".i_rm_n").displayed?
        number-=1
        browser.wait_to_appear(:id,"mpf0_MsgContainer")
      end
    ensure
      browser.close
      return @number_found
    end

  end

  def search(args)
    subject = args[:subject]
    from = args[:from]
    body = args[:body] 
    date = "1-Apr-1995"
    date = args[:date] if args[:date] #format of 1-Apr-2003"
    seen = ""
    seen = args[:seen] if args[:seen] # "SEEN" or "UNSEEN"
    number = -1
    number = args[:number] if args[:number]
    return search_imap(subject,from,body,number,date,seen) if @provider_info[:method] == "imap"          
    return [search_pop3(subject,from,body,number),search_owa(subject,from,body,number)] if @provider_info[:method] == "pop3"
  end

end
