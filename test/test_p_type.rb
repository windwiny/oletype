require "nokogiri"
require_relative "../downapi"

lines = DATA.read.lines

class MM < DownAPI
  def xx(fn, v)
    d = File.read(fn)
    hh = Nokogiri::HTML d
    parse_property_html("a", hh, fn, {})

    @properties.each do |k, kvs|
      if kvs["p_type"] == v
        puts "SAME #{v}                #{fn}"
      else
        p [kvs["p_type"], v, "  <<==========  #{fn} ", kvs["p_type_doc"], kvs["p_property_value_doc"], kvs["p_doc"], kvs["p_syntax_doc"], kvs["p_return_doc"]]
        puts
      end
    end
  end
end

lines.each do |line|
  line.strip!
  next if line[0] == "#" || line == ""
  vs = line.split(" ", 2)
  next unless vs.size > 0

  fn, v = vs[0], vs[1]
  if v && (v[0] == "\"" || v[0] == "\'")
    v = v.gsub('"', "").gsub("'", "")
  end

  m = MM.new
  m.xx fn, v
end

__END__
E:/mydata/pypp/oletype/exceldoc/excel.chart.hastitle.html bool
E:/mydata/pypp/oletype/exceldoc/excel.application.chartdatapointtrack.html bool
E:/mydata/pypp/oletype/exceldoc/excel.application.arbitraryxmlsupportavailable.html bool
E:/mydata/pypp/oletype/exceldoc/excel.application.calculationinterruptkey.html  XlCalculationInterruptKey
E:/mydata/pypp/oletype/exceldoc/excel.application.activecell.html Range
E:/mydata/pypp/oletype/exceldoc/excel.range.value.html  VBA_Variant
E:/mydata/pypp/oletype/exceldoc/excel.application.activesheet.html  "VBA_object | None"
E:/mydata/pypp/oletype/exceldoc/excel.application.sheetsinnewworkbook.html int
E:/mydata/pypp/oletype/exceldoc/excel.application.cursormovement.html "xlVisualCursor | xlLogicalCursor | int"
E:/mydata/pypp/oletype/exceldoc/excel.application.cutcopymode.html 'bool | XLCutCopyMode | int'
E:/mydata/pypp/oletype/exceldoc/excel.application.defaultsaveformat.html  "FileFormat | int"
E:/mydata/pypp/oletype/exceldoc/excel.workbook.fileformat.html XlFileFormat
E:/mydata/pypp/oletype/exceldoc/excel.application.mailsystem.html XlMailSystem
E:/mydata/pypp/oletype/exceldoc/excel.application.parent.html  VBA_object
E:/mydata/pypp/oletype/exceldoc/excel.application.height.html  float
E:/mydata/pypp/oletype/exceldoc/excel.addins.item.html  VBA_object
E:/mydata/pypp/oletype/exceldoc/excel.autocorrect.replacementlist.html list
E:/mydata/pypp/oletype/exceldoc/excel.pagesetup.evenpage.html PageSetup
E:/mydata/pypp/oletype/exceldoc/excel.application.enableanimations.html  __DEPRECATED_WARNNING__
E:/mydata/pypp/oletype/exceldoc/excel.fillformat.pictureeffects.html  VBA_object
E:/mydata/pypp/oletype/exceldoc/excel.application.replaceformat.html Replace
E:/mydata/pypp/oletype/exceldoc/excel.application.printcommunication.html bool
