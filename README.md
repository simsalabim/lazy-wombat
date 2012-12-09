# Lazy::Wombat ![Build Status](https://travis-ci.org/simsalabim/lazy-wombat.png "Build Status")

A simple yet powerful DSL to generate Excel spreadsheets built on top of [axlsx](https://github.com/randym/axlsx) gem.

## Why Lazy Wombat?
Axlsx is awesome and quite complex in usage.
Often you need something simple and easy-to-use to generate an excel spreadsheet as easy, as you markup tables with HTML.
Now you can.

## Installation

Add this line to your application's Gemfile:

    gem 'lazy-wombat'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install lazy-wombat

## Usage

### Basic sample: no styles, just content

Code
```ruby
LazyWombat::Xlsx.new do
  spreadsheet do
    row do
      cell 'Cyberdyne Systems'
      cell 'Model 101'
      cell 'The Terminator'
    end
  end
end.save 'my_laziness.xlsx'
```
Or shortened:
```ruby
LazyWombat::Xlsx.new do
  cell 'Cyberdyne Systems'
  cell 'Model 101'
  cell 'The Terminator'
end.save 'my_laziness.xlsx'
```
will create `my_laziness` spreadsheet looks like this: ![Generated spreadsheet](http://img525.imageshack.us/img525/7037/spreadsheet1.png)

Since spreadsheet elements inheritance is alike `spreadsheet -> row -> cell`, you can arbitrary omit every unnecessary
elder element of your spreadsheets.

### Where is my HTML??
Oh yeah, I promised you `as you markup tables with HTML`: there're logic aliases:
`spreadsheet` is `table`, `row` is `tr`, `cell` is, of course, `td`

Thus here's our simplest example:
```ruby
LazyWombat::Xlsx.new do
  table do
    tr do
      td 'Cyberdyne Systems'
      td 'Model 101'
      td 'The Terminator'
    end
  end
end.save 'my_laziness.xlsx'

# or
LazyWombat::Xlsx.new do
  td 'Cyberdyne Systems'
  td 'Model 101'
  td 'The Terminator'
end.save 'my_laziness.xlsx'
```

### Additional options: spreadsheet names, rows and cells styles
By default spreadsheets are named as `Sheet 1`, `Sheet 2`, etc. It can be overwritten using `name` option.
``ruby
LazyWombat::Xlsx.new do
  spreadsheet name: 'My Laziness' do
    row style: :bold do
      cell 'Cyberdyne Systems'
      cell 'Model 101', style: :italic
      cell 'The Terminator'
    end
  end
end.save 'my_laziness.xlsx'
```
![Generated spreadsheet](http://img521.imageshack.us/img521/4272/spreadsheet2.png)


## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
