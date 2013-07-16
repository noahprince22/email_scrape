require "bundler"
Bundler.require(:default)
#Provides added functionality to Selenium WebDriver
class TelluriumDriver

  def initialize(*args)
    @wait = Selenium::WebDriver::Wait.new(:timeout=>20)
     @driver = Selenium::WebDriver.for :firefox
  end
  
  def driver
    @driver
  end

  def go_to(url)
    current_name = driver.title
    driver.get url

    @wait.until { driver.title != current_name }
  end
  
  def load_website_and_wait_for_element(url,arg1,id)#takes a url, the symbol of what you wait to exist and the id/name/xpath, whatever of what you want to wait to exist
    current_name = driver.title
    driver.get url

    @wait.until { driver.title != current_name and driver.find_elements(arg1, id).size > 0 }
  end

  #takes the id you want to click, the id you want to change, and the value of the id you want
  # to change. (Example, I click start-application and wait for the next dialogue
  # box to not be hidden)
  def click_and_wait_to_change(id_to_click,id_to_change,value) 
    element_to_click = driver.find_element(:id, id_to_click)
    element_to_change = driver.find_element(:id, id_to_change)
    current_value = element_to_change.attribute(value.to_sym)

    element_to_click.click
    @wait.until { element_to_change.attribute(value.to_sym) != current_value }
  end

  def click_selector(id_to_click,selector_value)
    option = Selenium::WebDriver::Support::Select.new(driver.find_element(:id, id_to_click))
    option.select_by(:text, selector_value)
  end

  def click_selector_and_wait(id_to_click,selector_value,id_to_change,value)
    element_to_change = driver.find_element(:id => id_to_change)
    current_value = element_to_change.attribute(value.to_sym)

    option = Selenium::WebDriver::Support::Select.new(driver.find_element(:id => id_to_click))
    option.select_by(:text, selector_value)
      
    @wait.until { element_to_change.attribute(value.to_sym) != current_value }
  end

  def wait_for_element(element)
      i = 0
      @wait.until {
          bool = false

           if(element.displayed?)
              bool = true
              @element = element
              break
           end       

         bool == true
      }
  end

  def wait_to_disappear(sym,id)
   @wait.until {
    element_arr = driver.find_elements(sym,id)
    element_arr.size > 0 and !element_arr[0].displayed? #wait until the element both exists and is displayed
  }
  end
  
  def wait_to_appear(sym,id)
   @wait.until {
    element_arr = driver.find_elements(sym,id)
    element_arr.size > 0 and element_arr[0].displayed? #wait until the element both exists and is displayed
  }
  end
  
  def wait_and_click(sym, id)
   found_element = false

    #if the specified symbol and id point to multiple elements, we want to find the first visible one
    #and click that
    @wait.until do
      driver.find_elements(sym,id).shuffle.each do |element|
        if element.displayed? and element.enabled?
          @element=element
          found_element = true
        end 
      
      end
     found_element   
    end

    i = 0

    begin
      @element.click
    rescue Exception => ex
      i+=1
      sleep(1)
      retry if i<20
      raise ex 
    end
    
  end
 
  def wait_for_element_and_click(element)
    wait_for_element(element)
    
    i = 0
    begin
      @element.click
    rescue Exception => ex
      i+=1
      sleep(1)
      retry if i<20
      raise ex 
    end
    
  end

#hovers where the element is and clicks. Workaround for timeout on click that happened with the new update to the showings app
  def hover_click(element)
    driver.action.click(element).perform
  end

  def send_keys(sym,id,value)
    self.wait_and_click(sym, id)
    driver.action.send_keys(driver.find_element(sym => id),value).perform
  end

  #takes the id you want to click, the id you want to exist.(Example, I click start-application and wait for the next dialogue box to exist
  def click_and_wait_to_exist(*args) 
     case args.size
     when 2 #takes just two id's
       id1,id2 = args
        element_to_click = driver.find_element(:id => id1)
        element_to_change = driver.find_element(:id => id2)

        element_to_click.click
        wait_for_element(element_to_change)

     when 4 #takes sym,sym2,id1,id2
       sym,sym2,id1,id2 = args
        element_to_click = driver.find_element(sym => id)
        element_to_change = driver.find_element(sym2 => id2)

        element_to_click.click
        wait_for_element(element_to_change)

     else
        error
     end
  end

  def form_fillout(hash) #takes a hash of ID's with values to fillout
    hash.each do |id,value|
       self.send_keys(:id,id,value)
    end

  end
  
  def form_fillout_selector(hash)
    hash.each do |id,value|
      option = Selenium::WebDriver::Support::Select.new(driver.find_element(:id => id))
      option.select_by(:text,value)
    end

  end
  
  def wait_and_hover_click(sym,id)
    found = false
  #wait until an element with sym,id is displayed. When it is, hover click it
    @wait.until {
      elements = driver.find_elements(sym, id)
      elements.each do |element| 
        if element.displayed?
          found = true
          @element = element
        end
    
       end
      found == true
    } 
    self.hover_click(@element)
  end
  
  def close
    driver.quit
  end

end 
