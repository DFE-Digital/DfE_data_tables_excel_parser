# frozen_string_literal: true

source 'https://rubygems.org'

git_source(:github) { |repo_name| "https://github.com/#{repo_name}" }

gem 'roo', '~> 2.7'

# TODO: update to origin once https://github.com/piotrmurach/tty/pull/51 is merged
gem 'tty', github: 'spikeheap/tty', branch: 'patch-1'

gem "elasticsearch", "~> 6.1"


group :test, :development do
  gem "pry", "~> 0.12.2"
  gem "rubocop", "~> 0.64.0"
end

group :test do
  gem "rspec", "~> 3.8"

  gem "factory_bot", "~> 5.0"
  gem "faker", "~> 1.9"
end
