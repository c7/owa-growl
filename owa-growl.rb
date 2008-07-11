#!/usr/bin/env ruby

require 'rubygems'
require 'daemons'
require 'highline/import'
require 'mechanize'
require 'ruby-growl'

# Configuration
GROWL_HOST = 'localhost'
GROWL_NOTIFICATION_TITLE = 'New Email in Outlook Web Access'
FOLDERS_URL = 'YOUR_URL_TO_THE_FOLDERS_PAGE_IN_OWA'
SLEEP_DURATION = 300

if ARGV[0] == "start"
  # Get the required user input before we start the loop
  GROWL_PASSWORD = ask('Growl password: ') { |q| q.echo = "*" }
  OWA_USERNAME = ask('Email username: ')
  OWA_PASSWORD = ask('Email password: ') { |q| q.echo = "*" }
end

Daemons.run_proc('owa-growl') do
  loop do
    agent = WWW::Mechanize.new
    agent.auth(OWA_USERNAME, OWA_PASSWORD)
    page = agent.get(FOLDERS_URL)
    
    new_email_count = 0
    
    elements = page.search("/html/body//i")
    
    unless elements.empty?
      elements.each do |element|
        match = /\((\d+)\)/.match(element.inner_html)
        unless match.nil?
          new_email_count += match[1].to_i
        end
      end
    end
    
    if new_email_count > 0
      g = Growl.new GROWL_HOST, 'owa-growl', ["owa"], nil, GROWL_PASSWORD
      g.notify "owa", GROWL_NOTIFICATION_TITLE, "Email count: #{new_email_count}"
    end
    
    sleep SLEEP_DURATION
  end
end