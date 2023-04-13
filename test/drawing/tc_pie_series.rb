require 'tc_helper.rb'

class TestPieSeries < Test::Unit::TestCase
  def setup
    p = Axlsx::Package.new
    @ws = p.workbook.add_worksheet :name => "hmmm"
    chart = @ws.add_chart Axlsx::Pie3DChart, :title => "fishery"
    @series = chart.add_series :data => [0, 1, 2], :labels => ["zero", "one", "two"], :title => "bob", :colors => ["FF0000", "00FF00", "0000FF"]
  end

  def test_initialize
    assert_equal("bob", @series.title.text, "series title has been applied")
    assert_equal(@series.labels.class, Axlsx::AxDataSource)
    assert_equal(@series.data.class, Axlsx::NumDataSource)
    assert_nil(@series.explosion, "series shape has been applied")
  end

  def test_explosion
    assert_raise(ArgumentError, "require valid explosion") { @series.explosion = :lots }
    assert_nothing_raised("allow valid explosion") { @series.explosion = 20 }
    assert_equal(20, @series.explosion)
    # issue 58 - explosion caused to_xml_string to fail - now tested
    assert_nothing_raised("allow to_xml_string") { @series.to_xml_string }
  end

  def test_to_xml_string
    doc = Nokogiri::XML(@series.to_xml_string)

    assert(doc.xpath("//srgbClr[@val='#{@series.colors[0]}']"))
  end
  # TODO: test unique serialization parts
end
