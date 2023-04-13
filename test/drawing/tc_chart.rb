$LOAD_PATH.unshift "#{File.dirname(__FILE__)}/../"

require 'tc_helper'

class TestChart < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::Bar3DChart, :title => "fishery", :bg_color => "000000"
  end

  def teardown; end

  def test_initialization
    assert_equal(@p.workbook.charts.last, @chart, "the chart is in the workbook")
    assert_equal("fishery", @chart.title.text, "the title option has been applied")
    assert((@chart.series.is_a?(Axlsx::SimpleTypedList) && @chart.series.empty?), "The series is initialized and empty")
  end

  def test_title
    @chart.title.text = 'wowzer'

    assert_equal("wowzer", @chart.title.text, "the title text via a string")
    assert_nil(@chart.title.cell, "the title cell is nil as we set the title with text.")
    @chart.title = @row.cells.first

    assert_equal("one", @chart.title.text, "the title text was set via cell reference")
    assert_equal(@chart.title.cell, @row.cells.first)
    @chart.title = ""

    assert_empty(@chart.title)
  end

  def test_style
    assert_raise(ArgumentError) { @chart.style = 49 }
    assert_nothing_raised { @chart.style = 2 }
    assert_equal(2, @chart.style)
  end

  def test_to_from_marker_access
    assert(@chart.to.is_a?(Axlsx::Marker))
    assert(@chart.from.is_a?(Axlsx::Marker))
  end

  def test_bg_color
    assert_raise(ArgumentError) { @chart.bg_color = 2 }
    assert_nothing_raised { @chart.bg_color = "FFFFFF" }
    assert_equal("FFFFFF", @chart.bg_color)
  end

  def test_title_size
    assert_raise(ArgumentError) { @chart.title_size = 2 }
    assert_nothing_raised { @chart.title_size = "100" }
    assert_equal("100", @chart.title.text_size)
  end

  def test_vary_colors
    assert(@chart.vary_colors)
    assert_raise(ArgumentError) { @chart.vary_colors = 7 }
    assert_nothing_raised { @chart.vary_colors = false }
    refute(@chart.vary_colors)
  end

  def test_display_blanks_as
    assert_equal(:gap, @chart.display_blanks_as, "default is not :gap")
    assert_raise(ArgumentError, "did not validate possible values") { @chart.display_blanks_as = :hole }
    assert_nothing_raised { @chart.display_blanks_as = :zero }
    assert_nothing_raised { @chart.display_blanks_as = :span }
    assert_equal(:span, @chart.display_blanks_as)
  end

  def test_start_at
    @chart.start_at 15, 25

    assert_equal(15, @chart.graphic_frame.anchor.from.col)
    assert_equal(25, @chart.graphic_frame.anchor.from.row)
    @chart.start_at @row.cells.first

    assert_equal(0, @chart.graphic_frame.anchor.from.col)
    assert_equal(0, @chart.graphic_frame.anchor.from.row)
    @chart.start_at [5, 6]

    assert_equal(5, @chart.graphic_frame.anchor.from.col)
    assert_equal(6, @chart.graphic_frame.anchor.from.row)
  end

  def test_end_at
    @chart.end_at 25, 90

    assert_equal(25, @chart.graphic_frame.anchor.to.col)
    assert_equal(90, @chart.graphic_frame.anchor.to.row)
    @chart.end_at @row.cells.last

    assert_equal(2, @chart.graphic_frame.anchor.to.col)
    assert_equal(0, @chart.graphic_frame.anchor.to.row)
    @chart.end_at [10, 11]

    assert_equal(10, @chart.graphic_frame.anchor.to.col)
    assert_equal(11, @chart.graphic_frame.anchor.to.row)
  end

  def test_add_series
    s = @chart.add_series :data => [0, 1, 2, 3], :labels => ["one", 1, "anything"], :title => "bob"

    assert_equal(@chart.series.last, s, "series has been added to chart series collection")
    assert_equal("bob", s.title.text, "series title has been applied")
  end

  def test_pn
    assert_equal("charts/chart1.xml", @chart.pn)
  end

  def test_d_lbls
    assert_nil(Axlsx.instance_values_for(@chart)[:d_lbls])
    @chart.d_lbls.d_lbl_pos = :t

    assert(@chart.d_lbls.is_a?(Axlsx::DLbls), 'DLbls instantiated on access')
  end

  def test_plot_visible_only
    assert(@chart.plot_visible_only, "default should be true")
    @chart.plot_visible_only = false

    assert_false(@chart.plot_visible_only)
    assert_raise(ArgumentError) { @chart.plot_visible_only = "" }
  end

  def test_rounded_corners
    assert(@chart.rounded_corners, "default should be true")
    @chart.rounded_corners = false

    assert_false(@chart.rounded_corners)
    assert_raise(ArgumentError) { @chart.rounded_corners = "" }
  end

  def test_to_xml_string
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = schema.validate(doc).map { |error| puts error.message; error }

    assert_empty(errors, "error free validation")
  end

  def test_to_xml_string_for_display_blanks_as
    @chart.display_blanks_as = :span
    doc = Nokogiri::XML(@chart.to_xml_string)

    assert_equal("span", doc.xpath("//c:dispBlanksAs").attr("val").value, "did not use the display_blanks_as configuration")
  end

  def test_to_xml_string_for_title
    @chart.title = "foobar"
    doc = Nokogiri::XML(@chart.to_xml_string)

    assert_equal("foobar", doc.xpath("//c:title//c:tx//a:t").text)

    @chart.title = ""
    doc = Nokogiri::XML(@chart.to_xml_string)

    assert_equal(0, doc.xpath("//c:title").size)
  end

  def test_to_xml_string_for_plot_visible_only
    assert_equal("true", Nokogiri::XML(@chart.to_xml_string).xpath("//c:plotVisOnly").attr("val").value)
    @chart.plot_visible_only = false

    assert_equal("false", Nokogiri::XML(@chart.to_xml_string).xpath("//c:plotVisOnly").attr("val").value)
  end

  def test_to_xml_string_for_rounded_corners
    assert_equal("true", Nokogiri::XML(@chart.to_xml_string).xpath("//c:roundedCorners").attr("val").value)
    @chart.rounded_corners = false

    assert_equal("false", Nokogiri::XML(@chart.to_xml_string).xpath("//c:roundedCorners").attr("val").value)
  end
end
