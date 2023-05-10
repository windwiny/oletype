require "nokogiri"
require_relative "../downapi"

lines = DATA.read.lines

class MM < DownAPI
  def xx(fn, v)
    d = File.read(fn)
    hh = Nokogiri::HTML d
    parse_method_html("a", hh, fn, {})

    @methods.each do |k, kvs|
      if kvs["m_return"] == v
        puts "SAME #{v}                #{fn}"
      else
        p [kvs["m_return"], v, "  <<==========  #{fn} ", kvs["m_return_doc"], kvs["m_doc"]]
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
E:/mydata/pypp/oletype/exceldoc/excel.application.checkspelling.html bool
E:/mydata/pypp/oletype/exceldoc/excel.application.activatemicrosoftapp.html None
E:/mydata/pypp/oletype/exceldoc/excel.application.displayxmlsourcepane.html  None
E:/mydata/pypp/oletype/exceldoc/excel.range.flashfill.html  None
E:/mydata/pypp/oletype/exceldoc/excel.sheets.add.html VBA_object
E:/mydata/pypp/oletype/exceldoc/excel.groupshapes.item.html Shape
E:/mydata/pypp/oletype/exceldoc/excel.formatconditions.item.html VBA_object
