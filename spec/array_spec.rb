# encoding: utf-8
require File.expand_path(File.dirname(__FILE__) + '/spec_helper')

describe Array do

  describe :to_xls do
    it "should throw no error without data" do
      lambda { [].to_xls }.should_not raise_error
    end

=begin
    it "should return the correct string representing a spreadsheet" do
      s = StringIO.new
      xls = make_book(mock_users)
      xls.write(s)
      s.string.bytes.to_a.should == mock_users.to_xls.bytes.to_a
    end
=end
  end

end
