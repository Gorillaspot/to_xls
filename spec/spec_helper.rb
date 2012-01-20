$LOAD_PATH.unshift(File.join(File.dirname(__FILE__), '..', 'lib'))
$LOAD_PATH.unshift(File.dirname(__FILE__))
require 'rspec'
require 'factory_girl'
require 'to_xls'
require 'spreadsheet'

# Requires supporting files with custom matchers and macros, etc,
# in ./support/ and its subdirectories.
Dir["#{File.dirname(__FILE__)}/support/**.rb", ].each {|f| require f}

RSpec.configure do |config|
  
end

def mock_model(name, attributes)
  attributes[:attributes] = attributes.clone
  mock(name, attributes)
end

def mock_company(name, address)
  mock_model( name, :name => name, :address => address )
end

def mock_user(name, age, email, company)
  user = mock_model(name, :name => name, :age => age, :email => email)
  user.stub!(:company).and_return(company)
  user
end

def mock_users
  acme = mock_company('Acme', 'One Road')
  eads = mock_company('EADS', 'Another Road')

  [ mock_user('Peter', 20, 'peter@gmail.com', acme),
    mock_user('John',  25, 'john@gmail.com', acme),
    mock_user('Day9',  27, 'day9@day9tv.com', eads)
  ]
end

def check_sheet(sheet, array)
  sheet.rows.each_with_index do |row, i|
    row.should == array[i]
  end
end

def make_book(array, options={})
  book = Spreadsheet::Workbook.new
  ToXls::ArrayWriter.new(array, options).write_book(book)
  book
end
