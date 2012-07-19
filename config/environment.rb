#-*- coding: utf-8 -*-
require "rubygems"

# Set up gems listed in the Gemfile.
ENV['BUNDLE_GEMFILE'] ||= File.expand_path('./Gemfile')
require 'bundler/setup' if File.exists?(ENV['BUNDLE_GEMFILE'])

# other requires
require "sinatra/base"
require "sinatra/namespace"
require "json"

require "base64"
require "tempfile"

# defined application class
class ExcelServer < Sinatra::Base
  register Sinatra::Namespace
  # initialize
end

# require from ./app
Dir.chdir(File.expand_path("./app")) do |path|
  Dir.glob("*.rb").each do |file|
    require "#{path}/#{file}"
  end
end
