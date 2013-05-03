# Electricity Build Rate constraint

This is a quick attempt. It is probably both wrong and buggy, so don't trust it yet.

The purpose of this application is to help people to explore what would happen to electricity generation emissions in the UK out to 2050 under a variety of different assumptions. It is based around a very simple excel model.

## Canonical source

<https://github.com/decc/electricity-build-rate-constraint>

Report bugs to <https://github.com/decc/electricity-build-rate-constraint/issues>

## Dependencies

1. Ruby 2.0
2. A working C compiler
3. A webserver configured to run rack

To configure a webserver on OSX (done after setup below)

  gem install powder
  powder install
  cd /directory_where_this_program_is_installed
  powder link
  powder open
  
The resulting website only works properly on the latest generation of browsers.

Only tested on OSX 10.8.3 so far.

## Setup

  bundle
  bundle exec ruby make-model.rb

## Hacking

Written in a combination of Ruby and Coffeescript. 

  electricity-build-rate-constraint.xlsx - The actual model of the electricity system that we are exploring
  model.c - An computer-generated translation of the excel file into C
  model.rb - A ruby interface to model.c
  config.ru and server.rb - Serves 
    the data from model.rb
    public/index.html
    the javascripts refered to in src/javascripts/application.js
    the stylesheets refered to in src/stylesheets/application.css
  chart.js.coffee - The main program at the moment.

Fixes accepted, preferably using 'Fork&Pull' on github: <http://gun.io/blog/how-to-github-fork-branch-and-pull-request/>

## Licence

Open source: The MIT License (MIT)

Copyright (c) 2013 Tom Counsell

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN
THE SOFTWARE.
