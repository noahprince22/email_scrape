email_scrape
============
A simple library to search a given email address in the inbox and spambox for a subject, from, and body. 

Usage
-----------

        mail = Email.new(<username>@<provider>.com,<password>)
        mail.search({:body=> <body>, :subject => <subject>, :from => <from>, :seen = <SEEN, or UNSEEN>, :date => <date>, :number => <number to search>,:end_when_found => <true,false>})

You can search any, all, or none of those terms to find out how many emails exist with the given params in both the inbox and spambox
