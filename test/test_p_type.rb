require "nokogiri"
require_relative "../downapi"

lines = DATA.read.lines

class MM < DownAPI
  def xx(fn, v)
    d = File.read(fn)
    hh = Nokogiri::HTML d
    parse_property_html("a", hh, fn)

    @properties.each do |k, kvs|
      p [kvs["p_type"] == v, kvs["p_type"], kvs["p_type_doc"], kvs["p_property_value_doc"], kvs["p_doc"], kvs["p_syntax_doc"], kvs["p_return_doc"]]
    end
    puts
  end
end

lines.each do |line|
  line.strip!
  next if line[0] == "#" || line == ""
  vs = line.split()
  next unless vs.size > 0
  fn, v = vs[0], vs[1]
  m = MM.new
  m.xx fn, v
end

__END__
E:/mydata/pypp/oletype/exceldoc/excel.chart.hastitle.html bool
E:/mydata/pypp/oletype/exceldoc/excel.application.chartdatapointtrack.html bool
E:/mydata/pypp/oletype/exceldoc/excel.application.arbitraryxmlsupportavailable.html bool
E:/mydata/pypp/oletype/exceldoc/excel.application.calculationinterruptkey.html  XlCalculationInterruptKey
E:/mydata/pypp/oletype/exceldoc/excel.application.activecell.html Range
E:/mydata/pypp/oletype/exceldoc/excel.application.enableanimations.html  __UNKNOWN_TYPE__
E:/mydata/pypp/oletype/exceldoc/excel.range.value.html  VBA_Variant
E:/mydata/pypp/oletype/exceldoc/excel.application.activesheet.html  __UNKNOWN_TYPE__
E:/mydata/pypp/oletype/exceldoc/excel.application.sheetsinnewworkbook.html int
