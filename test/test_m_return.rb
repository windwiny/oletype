require "nokogiri"
require_relative "../downapi"

lines = DATA.read.lines

class MM < DownAPI
  def xx(fn, v)
    d = File.read(fn)
    hh = Nokogiri::HTML d
    parse_method_html("a", hh, fn)

    @methods.each do |k, kvs|
      p [kvs["m_return"] == v, kvs["m_return"], kvs["m_return_doc"], kvs["m_doc"]]
    end
  end
end

lines.each do |line|
  vs = line.split()
  fn, v = vs[0], vs[1]
  m = MM.new
  m.xx fn, v
end

__END__
E:/mydata/pypp/oletype/exceldoc/excel.application.checkspelling.html bool
E:/mydata/pypp/oletype/exceldoc/excel.application.activatemicrosoftapp.html
E:/mydata/pypp/oletype/exceldoc/excel.application.displayxmlsourcepane.html
E:/mydata/pypp/oletype/exceldoc/excel.range.flashfill.html  None
