# -*- encoding: utf-8 -*-
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'lazy-wombat/version'

Gem::Specification.new do |s|
  s.name          = 'lazy-wombat'
  s.version       = LazyWombat::VERSION
  s.authors       = ['Alexander Kaupanin']
  s.email         = %w(kaupanin@gmail.com)
  s.description   = %q{A simple yet powerful DSL to create Excel spreadsheets built on top of axlsx gem}
  s.summary       = %q{A simple yet powerful DSL to create Excel spreadsheets built on top of axlsx gem}
  s.homepage      = 'http://github.com/simsalabim/lazy-wombat'

  s.files         = `git ls-files`.split($/)
  s.executables   = s.files.grep(%r{^bin/}).map{ |f| File.basename(f) }
  s.test_files    = s.files.grep(%r{^(test|spec|features)/})
  s.require_paths = %w(lib)

  s.add_development_dependency 'cucumber', '~> 1.1'
  s.add_development_dependency 'rspec',    '~> 2.9'
  s.add_development_dependency 'roo'
  s.add_runtime_dependency     'axlsx'
end
