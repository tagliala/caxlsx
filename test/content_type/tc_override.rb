require 'tc_helper'

class TestOverride < Test::Unit::TestCase
  def test_content_type_restriction
    assert_raise(ArgumentError, "requires known content type") { Axlsx::Override.new :ContentType => "asdf" }
  end

  def test_to_xml
    type = Axlsx::Override.new :PartName => "somechart.xml", :ContentType => Axlsx::CHART_CT
    doc = Nokogiri::XML(type.to_xml_string)

    assert_equal(1, doc.xpath("Override[@ContentType='#{Axlsx::CHART_CT}']").size)
    assert_equal(1, doc.xpath("Override[@PartName='somechart.xml']").size)
  end
end
