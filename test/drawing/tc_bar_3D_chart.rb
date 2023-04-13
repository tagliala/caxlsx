require 'tc_helper.rb'

class TestBar3DChart < Test::Unit::TestCase
  def setup
    @p = Axlsx::Package.new
    ws = @p.workbook.add_worksheet
    @row = ws.add_row ["one", 1, Time.now]
    @chart = ws.add_chart Axlsx::Bar3DChart, :title => "fishery"
  end

  def teardown; end

  def test_initialization
    assert_equal(:clustered, @chart.grouping, "grouping defualt incorrect")
    assert_equal(@chart.series_type, Axlsx::BarSeries, "series type incorrect")
    assert_equal(:bar, @chart.bar_dir, " bar direction incorrect")
    assert(@chart.cat_axis.is_a?(Axlsx::CatAxis), "category axis not created")
    assert(@chart.val_axis.is_a?(Axlsx::ValAxis), "value access not created")
  end

  def test_bar_direction
    assert_raise(ArgumentError, "require valid bar direction") { @chart.bar_dir = :left }
    assert_nothing_raised("allow valid bar direction") { @chart.bar_dir = :col }
    assert_equal(:col, @chart.bar_dir)
  end

  def test_grouping
    assert_raise(ArgumentError, "require valid grouping") { @chart.grouping = :inverted }
    assert_nothing_raised("allow valid grouping") { @chart.grouping = :standard }
    assert_equal(:standard, @chart.grouping)
  end

  def test_gap_width
    assert_raise(ArgumentError, "require valid gap width") { @chart.gap_width = -1 }
    assert_raise(ArgumentError, "require valid gap width") { @chart.gap_width = 501 }
    assert_nothing_raised("allow valid gapWidth") { @chart.gap_width = 200 }
    assert_equal(200, @chart.gap_width, 'gap width is incorrect')
  end

  def test_gap_depth
    assert_raise(ArgumentError, "require valid gap_depth") { @chart.gap_depth = -1 }
    assert_raise(ArgumentError, "require valid gap_depth") { @chart.gap_depth = 501 }
    assert_nothing_raised("allow valid gap_depth") { @chart.gap_depth = 200 }
    assert_equal(200, @chart.gap_depth, 'gap depth is incorrect')
  end

  def test_shape
    assert_raise(ArgumentError, "require valid shape") { @chart.shape = :star }
    assert_nothing_raised("allow valid shape") { @chart.shape = :cone }
    assert_equal(:cone, @chart.shape)
  end

  def test_to_xml_string
    schema = Nokogiri::XML::Schema(File.open(Axlsx::DRAWING_XSD))
    doc = Nokogiri::XML(@chart.to_xml_string)
    errors = []
    schema.validate(doc).each do |error|
      errors.push error
      puts error.message
    end

    assert_empty(errors, "error free validation")
  end

  def test_to_xml_string_has_axes_in_correct_order
    str = @chart.to_xml_string
    cat_axis_position = str.index(@chart.axes[:cat_axis].id.to_s)
    val_axis_position = str.index(@chart.axes[:val_axis].id.to_s)

    assert(cat_axis_position < val_axis_position, "cat_axis must occur earlier than val_axis in the XML")
  end

  def test_to_xml_string_has_gap_depth
    gap_depth_value = rand(0..500)
    @chart.gap_depth = gap_depth_value
    doc = Nokogiri::XML(@chart.to_xml_string)

    assert_equal(doc.xpath("//c:bar3DChart/c:gapDepth").first.attribute('val').value, gap_depth_value.to_s)
  end

  def test_to_xml_string_has_gap_width
    gap_width_value = rand(0..500)
    @chart.gap_width = gap_width_value
    doc = Nokogiri::XML(@chart.to_xml_string)

    assert_equal(doc.xpath("//c:bar3DChart/c:gapWidth").first.attribute('val').value, gap_width_value.to_s)
  end
end
