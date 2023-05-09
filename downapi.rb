require "json"
require "set"
require "stringio"
require "net/http"
require "nokogiri"

# Config TODO  sync rb/py file
module Cfg
  OUT_EXCEL_INFO_FN = "excel.info.json"

  N_SUMM = "summary"

  N_co = "collections"
  N_co_doc = "co_doc"

  N_e = "enumerations"
  N_e_doc = "e_doc"
  N_e_table_head = "e_table_head"
  N_e_table_rows = "e_table_rows"
  N_e_remarks_doc = "e_remarks_doc"
  N_e_table_format_incorrect = "e_ERROR"

  N_c = "classes"
  N_c_doc = "c_doc"
  N_c_remarks_doc = "c_remarks_doc"
  N_c_example_doc = "c_examples_doc"

  N_m = "methods"
  N_m_doc = "m_doc"
  N_m_return = "m_return"
  N_m_return_doc = "m_return_doc"
  N_m_parameters_doc = "m_parameters_doc"
  N_m_remarks_doc = "m_remarks_doc"
  N_m_example_doc = "m_example_doc"

  N_p = "properties"
  N_p_doc = "p_doc"
  N_p_type = "p_type"
  N_p_syntax_doc = "p_syntax_doc"
  N_p_return_doc = "p_return_doc"
  N_p_property_value_doc = "p_property_value_doc"
  N_p_remarks_doc = "p_remarks_doc"
  N_p_example_doc = "p_example_doc"

  DOCDIR = "exceldoc"
  CACHEDIR = File.join(__dir__, DOCDIR)
end

include Cfg

# TODO  vba's type to python type
VBAtype2pytype = {
  "Boolean" => "bool",
  "Byte" => "VBA_Byte",
  "Currency" => "VBA_Currency",
  "Date" => "datetime.datetime",
  "Decimal" => "VBA_Decimal",
  "Double" => "float",
  "Integer" => "int",
  "Long" => "int",
  "LongLong" => "int",
  "LongPtr" => "int",
  "Object" => "VBA_Object",
  "Single" => "float",
  "String" => "str",
  "Variant" => "VBA_Variant",

  "OBJECT" => "VBA_OBJECT",
  "object" => "VBA_object",
  "BOOL" => "bool",
  "Bool" => "bool",
  "True" => "bool",
  "False" => "bool",
  "OK" => "bool",
  "Cancel" => "bool",

  "Nothing" => "Nothing",
  "VOID" => "VOID",
  "Void" => "Void",
}

# TODO
module OLE_APPNAME
  EXCEL = "(Excel)"
  OFFICE = "(Office)"
  POWERPOINT = "(PowerPoint)"
  WORD = "(Word)"
end

module OLE_TYPE
  COLLECTION = "Collection"
  ENUMERATION = "enumeration"
  OBJECT = "object"
  PROPERTY = "property"
  METHOD = "method"
  EVENT = "event"
end

def url_remove_dotdot(u1)
  #u1='https://learn.microsoft.com/en-us/office/vba/api/../excel/concepts/workbooks-and-worksheets/../../../api/excel.xmlnamespace.prefix'
  #u2='https://learn.microsoft.com/en-us/office/vba/api/excel.xmlnamespace.prefix'
  vs = u1.split("/")
  1.upto(vs.size).each do |i|
    if vs[i] == ".."
      vs[i] = nil
      j = i - 1
      while vs[j] == nil
        break if j < 0   #incorrect urlstr
        j = j - 1
      end
      vs[j] = nil
    end
  end
  u2 = vs.compact.join("/")
  u2
end

def uniq_type(ar)
  case ar[-1]
  when "Boolean"
    return "Boolean"
  when "Variant"
    return "Variant"
  when "Long"
    return "int"
  end
  res = ar.map do |e|
    e2 = VBAtype2pytype[e]
    if e2
      return e2
    elsif e.include?(" ")
      nil
    else
      e
    end
  end.compact.uniq
  if res.size > 0
    return res.join("_or_")
  else
    return nil
  end
end

class MyHash < Hash
  def []=(x, y)
    if self.has_key?(x)
      if self[x] != y
        raise "MyHASH ERROR has #{x}, \n v:#{y}, \nov:#{self[x]}"
      end
    end
    super
  end
end

class DownAPI
  def initialize(*urls)
    @links_all = Hash.new()
    urls.each { |url|
      @links_all[url] = ""
    }
    @links_skips = Set.new()
    @url_info_kvs = MyHash.new

    @parsed_html = 0
    @undown = []
    @downloaded = Set.new()

    @collection = MyHash.new
    @enumeration = MyHash.new
    @enumeration_table_not_found = 0

    @classes = MyHash.new
    @properties = MyHash.new
    @methods = MyHash.new

    if __FILE__ == $0
      at_exit { finish() }
      @th = Thread.new do
        ii = 0
        t0 = Time.now
        sum0 = @parsed_html
        pp1 = Proc.new do
          ii += 1
          t1 = Time.now
          t1s = t1.strftime("%H%M%S")
          msg = %{WATCH.. #{ii} #{t1s} speed:#{((@parsed_html - sum0) / (t1 - t0)).to_i}/s  links parsed:#{@parsed_html} all:#{@links_all.size} undown:#{@undown.size} downed:#{@downloaded.size} skips:#{@links_skips.size}  url_info_kvs:#{@url_info_kvs.size}\n}
          STDERR.print msg
          t0 = t1
          sum0 = @parsed_html
        end
        at_exit { pp1.call }
        while true
          sleep 2
          pp1.call
        end
      end
    end
  end

  def download_all()
    @undown = @links_all.to_a

    while true
      @undown.each do |url, objtype2|
        if @downloaded.include?(url)
          next
        end
        @downloaded.add(url)
        process_page(url, objtype2)
      end

      if @downloaded.to_a == @links_all.keys() # all links downed
        STDERR.puts %{ ALL DOWNLOADED, QUIT}
        break
      end
      cur_undown = @links_all.reject { |e| @downloaded.include?(e) }.to_a
      if cur_undown == @undown # some thing cannot download
        STDERR.puts %{ QUIT: still have those undownload link, break}
        @undown.each do |x|
          STDERR.puts %{  #{x}}
        end
        STDERR.puts
        break
      end

      @undown = cur_undown
    end
  end

  def finish()
    @links_all.each do |url, type|
      fn = url_fn(url)
      fn = %{#{DOCDIR}/#{File.basename(fn)}}
      @url_info_kvs[url] ||= {}
      @url_info_kvs[url]["htmlf"] = fn
      @url_info_kvs[url]["htmlf_file_exists"] = File.exist?(fn)
    end
    url_info_kvs_summ = @url_info_kvs.group_by { |url, kvs| kvs["type"] }.map { |ty, v| [ty.to_s, v.size] }.sort.to_h
    url_info_kvs_proc_summ = @url_info_kvs.group_by { |url, kvs| kvs["UN_PROCESS"] }.map { |b, v| [b.to_s, v.size] }.sort.to_h

    all_data = {
      "TODO" => "find ': null' check something not doc founded",
      N_SUMM => {
        downloads: @links_all.size,
        parsed: @classes.size + @methods.size + @properties.size,
        parse_skips: @url_info_kvs.select { |url, kvs| kvs.has_key?("UN_PROCESS") }.size,
        N_co => @collection.size,
        N_e => @enumeration.size,
        N_c => @classes.size,
        N_m => @methods.size,
        N_p => @properties.size,
        enumeration_table_not_found: @enumeration_table_not_found,
      },
      N_co => @collection,
      N_e => @enumeration,
      N_c => @classes,
      N_m => @methods,
      N_p => @properties,
      urls: {
        urls_summ: %{links all: #{@links_all.size} undown:#{@undown.size} skips:#{@links_skips.size} url_info_kvs:#{@url_info_kvs.size}},
        url_info_kvs_summ: url_info_kvs_summ,
        url_info_kvs_proc_summ: url_info_kvs_proc_summ,
        undown: @undown.sort,
        links_skips: @links_skips.to_a.sort,
        url_info_kvs: @url_info_kvs.sort_by { |url, kvs| [kvs["type"].to_s, kvs["title"].to_s] },
      },
    }

    if File.exist?(OUT_EXCEL_INFO_FN) && File.binread(OUT_EXCEL_INFO_FN) == all_data
      STDERR.puts(%{SKIP SAME OUTPUT files: #{OUT_EXCEL_INFO_FN}\n\n})
    else
      File.write(OUT_EXCEL_INFO_FN, JSON.pretty_generate(all_data))
      STDERR.puts(%{OUTPUT files: #{OUT_EXCEL_INFO_FN}\n\n})
    end
  end

  protected

  def process_page(url, objtype2)
    txt = download_page(url)
    @url_info_kvs[url] ||= {}

    hh = Nokogiri::HTML txt
    page_title = hh.css "div > h1"
    if page_title.size < 1
      @url_info_kvs[url]["UN_PROCESS"] = "todo NOT GET TITLE"
      return
    end
    app_name, objtype, obj_name = add_url_info(url, page_title[0].text)
    # TODO FIXME obj_name may has '.'

    if app_name != ::OLE_APPNAME::EXCEL
      @url_info_kvs[url]["UN_PROCESS"] = "todo NOT PROCESS APP"
      return
    end

    case objtype
    when ::OLE_TYPE::PROPERTY
      parse_property_html(url, hh, obj_name)
    when ::OLE_TYPE::METHOD
      parse_method_html(url, hh, obj_name)
    when ::OLE_TYPE::OBJECT
      es, ms, ps = parse_object_html(url, hh, obj_name)
      # TODO not need
      # es.map { |t,href| @links_all[href] = ::OLE_TYPE::EVENT }
      ps.map { |t, href| @links_all[href] = ::OLE_TYPE::PROPERTY }
      ms.map { |t, href| @links_all[href] = ::OLE_TYPE::METHOD }
    when ::OLE_TYPE::COLLECTION
      process_collection_html(url, hh, app_name, objtype, obj_name)
    when ::OLE_TYPE::ENUMERATION
      process_enumeration_html(url, hh, obj_name)
    else
      @url_info_kvs[url]["UN_PROCESS"] = "todo EXCEL OBJ OTHER"
    end
    @parsed_html += 1
  end

  protected

  def assert(x, msg)
    if !x
      raise %{#{msg} Failed}
    end
  end

  def url_split(url)
    url = url.to_s
    ix = url.rindex("/")
    if ix
      base = url[..ix]
      fn = url[ix + 1..]
    else
      base = ""
      fn = url
    end
    fn = fn.gsub(/[\\\/\:\*\?\"\<\>\|]/, "_") # illegal filename character on win \/:*?"<>|
    fn = "#{CACHEDIR}/#{fn}.html"
    [base, fn]
  end

  def url_base(url)
    base, _ = url_split(url)
    base
  end

  def url_fn(url)
    _, fn = url_split(url)
    fn
  end

  def baseurl_to_fullurl(baseurl, url)
    url = if url.include?("/") && url[0] != "."
        if url =~ /microsoft.com/
          url
        else
          @links_skips.add(url)
          nil
        end
      else
        baseurl + url
      end
    if url
      if url !~ /vbe-glossary/ || url =~ /vbe-glossary#point/ # TODO FIXME
        url_remove_dotdot(url)
      else
        @links_skips.add(url)
        nil
      end
    else
      nil
    end
  end

  def a2arr(baseurl, ess)
    res = ess.map do |e|
      href = baseurl_to_fullurl(baseurl, e["href"])
      if href
        [e.text, href]
      else
        nil
      end
    end
    res = res - [nil]
    pprintres(res)
    res
  end

  def pprintres(res)
    res.each do |x, y|
      if y == nil
        STDERR.puts %{    EEE URL  #{x} #{y}}
      end
    end
  end

  def save_html(url, txt)
    fn = url_fn(url)
    if !txt || txt.strip == ""
      STDERR.puts %{SKIP EMPTY #{url}} if $DEBUG
      return
    end
    STDERR.puts %{ SAVE #{File.basename(fn)},  #{txt.size}} if $DEBUG
    File.write(fn, txt)
  end

  def download_page(url)
    if !url
      STDERR.puts(%{?? EEE empty "#{url}" })
      return
    end
    fn = url_fn(url)
    if File.exist?(fn)
      STDERR.puts %{SKIP #{url} } if $DEBUG
      txt = File.read(fn)
      return txt
    end

    url = URI(url)

    STDERR.puts %{DOWNLOADING #{url} }
    begin
      r = Net::HTTP.get_response(url)
      txt = r.body
      save_html(url, txt)
      return txt
    rescue Exception => e
      STDERR.puts %{Net::HTTP ERROR #{e} on #{url}}
    end
    ""
  end

  def parse_event_html(url, hh, xx)
    raise "No implement yet"  # TODO FIXME
  end

  def foreach_a_add_to_links(retss, url)
    retss.each do |ret|
      aas = ret.css("a")
      aas.map { |a|
        aurl = a["href"]
        url2 = baseurl_to_fullurl(url_base(url), aurl)
        if url2
          @links_all[url2] = ::OLE_TYPE::OBJECT
        else
          @links_skips.add %{empty url:#{url}}
        end
        a.text.strip
      }
    end
  end

  def add_url_info(url, title)
    @url_info_kvs[url] ||= {}
    vs = title.split
    if vs[-1][0] == "("
      assert(vs.size >= 3, "title splie size >= 3 ERR. #{vs.size}  #{vs}")
      app_name = vs[-1]
      obj_type = vs[-2]
      obj_name = vs[0]
      ty = vs[-1] + " " + vs[-2]
    else
      app_name = title
      obj_type = title
      obj_name = title
      ty = "OTHER"
    end
    @url_info_kvs[url]["title"] = title
    @url_info_kvs[url]["type"] = ty

    [app_name, obj_type, obj_name]
  end

  def process_collection_html(url, hh, app_name, objtype, obj_name)
    if app_name == ::OLE_APPNAME::EXCEL && objtype == ::OLE_TYPE::COLLECTION
      obj_name = "VBA_" + ::OLE_TYPE::COLLECTION
    end
    @collection[obj_name] = iffo = { N_co_doc => nil }

    retss = hh.css 'nav[id="center-doc-outline"] ~ p,pre'
    if retss.size > 0
      iffo[N_co_doc] = retss.map { |e| e.text.strip }.join("\n\n")
      foreach_a_add_to_links(retss, url)
    end
  end

  def process_enumeration_html(url, hh, obj_name)
    @enumeration[obj_name] = iffo = { N_e_doc => nil }

    retss = hh.css 'nav[id="center-doc-outline"] + p'
    if retss.size > 0
      iffo[N_e_doc] = retss[0].text.strip
    end

    retss = hh.css 'nav[id="center-doc-outline"] +p +table'
    if retss.size > 0
      foreach_a_add_to_links(retss, url)
      ts = hh.css "table"
      tab1 = retss[0]

      th = tab1.css("th").map(&:text)
      if th[0..2] != %w{Name Value Description}
        iffo[N_e_table_format_incorrect] = true
      end
      rows = tab1.css("tr").map { |tr| tr.css("td").map(&:text) } - [[]]
      iffo[N_e_table_head] = th
      iffo[N_e_table_rows] = rows
    else
      iffo[N_e_table_head] = []
      iffo[N_e_table_rows] = []
      @enumeration_table_not_found += 1
    end

    retss = hh.css 'div > h2[id="remarks"] ~ p'
    if retss.size > 0
      iffo[N_e_remarks_doc] = retss.map { |e| e.text.strip }.join("\n\n")
      foreach_a_add_to_links(retss, url)
    end
  end

  def parse_property_html(url, hh, prop_name)
    @properties[prop_name] = iffo = { N_p_doc => nil, N_p_type => nil }

    retss = hh.css 'div > h2[id="return-value"] + p'
    if retss.size > 0
      foreach_a_add_to_links(retss, url)
      ty = retss[0].text.strip
      iffo[N_p_return_doc] = ty
      ty = find_property_type_name(retss, url)
      if ty
        iffo[N_p_type] = ty
      else
        iffo[N_p_type] = "__UN_PARSED_RETURN_VALUE__"
      end
    end

    retss = hh.css 'div > h2[id="property-value"] + p'
    if retss.size > 0
      foreach_a_add_to_links(retss, url)
      ty = retss[0].text.strip
      iffo[N_p_property_value_doc] = ty
      ty = find_property_type_name(retss, url)
      if ty
        iffo[N_p_type] = ty
      else
        iffo[N_p_type] = "__UN_PARSED_PROPERTY_VALUE__"
      end
    end

    retss = hh.css 'nav[id="center-doc-outline"] + p'
    if retss.size > 0
      iffo[N_p_doc] = retss[0].text.strip
      foreach_a_add_to_links(retss, url)
      ty = find_property_type_name(retss, url)
      iffo[N_p_type] = ty if !iffo[N_p_type] && ty
    else
      retss = hh.css 'nav[id="center-doc-outline"] + div'
      if retss.size > 0
        iffo[N_p_doc] = retss[0].text.strip
        ty = find_property_type_name(retss, url)
        iffo[N_p_type] = ty if !iffo[N_p_type] && ty
      end
    end

    retss = hh.css 'div > h2[id="syntax"] ~ p'
    if retss.size > 0
      iffo[N_p_syntax_doc] = retss.map { |e| e.text.strip }.join("\n\n")
      foreach_a_add_to_links(retss, url)
    end

    retss = hh.css 'div > h2[id="remarks"] ~ p'
    if retss.size > 0
      iffo[N_p_remarks_doc] = retss.map { |e| e.text.strip }.join("\n\n")
      foreach_a_add_to_links(retss, url)
    end

    retss = hh.css 'div > h2[id="example"] ~ p,pre'
    if retss.size > 0
      iffo[N_p_example_doc] = retss.map { |e| e.text.strip }.join("\n\n")
      foreach_a_add_to_links(retss, url)
    end

    iffo[N_p_type] = "__UNKNOWN_TYPE__" unless iffo[N_p_type]
  end

  def find_property_type_name(retss, url)
    # TODO FIXME

    retss.map do |res|
      ty = res.text.strip
      if VBAtype2pytype.has_key?(ty)
        return VBAtype2pytype[ty]
      elsif ty !~ /\s/
        return ty
      end

      if /by the collection/i =~ ty
        return nil  # collection
      end
    end

    tys = retss.map { |e| e.css "strong,a" }
      .flatten
      .map { |e| e.text }
      .reject { |e| e.include?(" ") }
      .map { |e| VBAtype2pytype.has_key?(e) ? VBAtype2pytype[e] : e }
      .uniq

    if tys.size == 0
      return nil
    elsif tys.size == 1
      ty = tys[0]
      if %w{Nothing VOID Void}.include?(ty)
        return nil
      end
      return ty
    else
      return nil
    end
  end

  def find_method_type_name(retss, url)
    # TODO FIXME
    retss.map do |res|
      ty = res.text.strip
      if VBAtype2pytype.has_key?(ty)
        return VBAtype2pytype[ty]
      elsif ty !~ /\s/
        return ty
      end

      if /by the collection/i =~ ty
        return nil  # collection
      end
    end

    tys = retss.map { |e| e.css "strong,a" }
      .flatten
      .map { |e| e.text }
      .reject { |e| e.include?(" ") }
      .map { |e| VBAtype2pytype.has_key?(e) ? VBAtype2pytype[e] : e }
      .uniq

    if tys.size == 0
      return nil
    elsif tys.size == 1
      return tys[0]
    else
      return tys.join("_or_")
    end
  end

  def parse_method_html(url, hh, method_name)
    @methods[method_name] = iffo = { N_m_doc => nil, N_m_return => nil }

    retss = hh.css 'div > h2[id="return-value"] + p'
    if retss.size > 0
      ty = find_method_type_name(retss, url)
      if ty
        iffo[N_m_return] = ty
      else
        iffo[N_m_return] = "__UN_PARSED_TYPE__" # un parsed string
      end
      iffo[N_m_return_doc] = retss.text.strip
      foreach_a_add_to_links(retss, url)
    end

    retss = hh.css 'nav[id="center-doc-outline"] + p'
    if retss.size > 0
      iffo[N_m_doc] = retss[0].text.strip
      unless iffo[N_m_return]
        ty = find_method_type_name(retss, url)
        iffo[N_m_return] = ty if ty
      end
      foreach_a_add_to_links(retss, url)
    end

    retss = hh.css 'div > h2[id="parameters"] + table'
    if retss.size > 0
      iffo[N_m_parameters_doc] = retss[0].text.gsub(/((\r)?\n){3,}/, "\n\n").gsub(/(?<!\n)\n(?!\n)/, " ").strip
    end

    retss = hh.css 'div > h2[id="remarks"] ~ p'
    if retss.size > 0
      iffo[N_m_remarks_doc] = retss.map { |e| e.text.strip }.join("\n\n")
      foreach_a_add_to_links(retss, url)
    end

    retss = hh.css 'div > h2[id="example"] ~ p,pre'
    if retss.size > 0
      iffo[N_m_example_doc] = retss.map { |e| e.text.strip }.join("\n\n")
      foreach_a_add_to_links(retss, url)
    end

    iffo.delete(N_m_return) if iffo[N_m_return] == nil
  end

  def parse_object_html(url, hh, cls_name)
    @classes[cls_name] = iffo = { N_c_doc => nil }

    retss = hh.css 'nav[id="center-doc-outline"] + p'
    if retss.size > 0
      iffo[N_c_doc] = retss[0].text.strip
      foreach_a_add_to_links(retss, url)
    end
    retss = hh.css 'div > h2[id="remarks"] ~ p'
    if retss.size > 0
      iffo[N_c_remarks_doc] = retss.map { |e| e.text.strip }.join("\n\n")
      foreach_a_add_to_links(retss, url)
    end

    retss = hh.css 'div > h2[id="example"] ~ p,pre'
    if retss.size > 0
      iffo[N_c_example_doc] = retss.map { |e| e.text.strip }.join("\n\n")
      foreach_a_add_to_links(retss, url)
    end

    events = hh.css 'div > h2[id="events"] + ul'
    assert(events.size <= 1, %{events ul > 1, #{events.size}  #{url}})
    methods = hh.css 'div > h2[id="methods"] + ul'
    assert(methods.size <= 1, %{methods ul > 1, #{methods.size}  #{url}})
    properties = hh.css 'div > h2[id="properties"] + ul'
    assert(properties.size <= 1, %{properties ul > 1, #{properties.size}  #{url}})

    baseurl = url_base(url)
    events = if events.size > 0
        a2arr(baseurl, events[0].css("a"))
      else
        []
      end
    methods = if methods.size > 0
        a2arr(baseurl, methods[0].css("a"))
      else
        []
      end
    properties = if properties.size > 0
        a2arr(baseurl, properties[0].css("a"))
      else
        []
      end
    STDERR.puts %{ HTML #{txt.size} bytes. SUMMARY: events #{events.size},  methods #{methods.size}, properties #{properties.size}} if $DEBUG

    [events, methods, properties]
  end
end

if __FILE__ == $0
  Dir.mkdir(CACHEDIR) unless File.exist?(CACHEDIR)

  begin_url = "https://learn.microsoft.com/en-us/office/vba/api/excel.application(object)"

  dd = DownAPI.new(begin_url)
  dd.download_all()
end
