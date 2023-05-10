require "json"
require "set"
require "stringio"
require "net/http"
require "nokogiri"

# Config TODO  sync rb/py file
module Cfg
  OUT_EXCEL_INFO_FN = "excel.info.json"

  N_SUMM = "summary"
  N_META = "meta"

  N_co = "collections"
  N_co_doc = "co_doc"

  N_e = "enumerations"
  N_e_unique = "uniqued"
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
  "Date" => "Date", #"datetime.datetime",
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
  "String-Variant" => "VBA_Variant",

  "OBJECT" => "VBA_Object",
  "object" => "VBA_Object",
  "BOOL" => "bool",
  "Bool" => "bool",
  "True" => "bool",
  "False" => "bool",
  "OK" => "bool",
  "Cancel" => "bool",
  "True/False" => "bool",
  "unique" => "UNIQUE",  #TODO

  "Nothing" => "None",
  "VOID" => "None",
  "Void" => "None",
  "array" => "list",

  "INT32" => "int",
  "constants" => "constants",
  "currency" => "currency",
  "date" => "Date",
  "general" => "general",
  "integer" => "int",
  "null" => "None",
  "points" => "points",
  "single" => "float",
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

# TODO html contents css path
module HTMLCSS
  H1_TITLE = "div > h1"

  ID_DOC_p = 'nav[id="center-doc-outline"] + p'
  ID_DOC_p_pres = 'nav[id="center-doc-outline"] ~ p,pre'
  ID_DOC_div = 'nav[id="center-doc-outline"] + div'
  ID_DOC_table = 'nav[id="center-doc-outline"] +p +table'
  ID_PROPERTY_VALUE_p = 'div > h2[id="property-value"] + p'
  ID_RETURN_VALUE_p = 'div > h2[id="return-value"] + p'

  ID_SYNTAX_ps = 'div > h2[id="syntax"] ~ p'
  ID_REMARKS_ps = 'div > h2[id="remarks"] ~ p'
  ID_EXAMPLE_p_pres = 'div > h2[id="example"] ~ p,pre'
  ID_PARAMETERS_table = 'div > h2[id="parameters"] + table'

  ID_EVENTS_ul = 'div > h2[id="events"] + ul'
  ID_METHODS_ul = 'div > h2[id="methods"] + ul'
  ID_PROPERTIES_ul = 'div > h2[id="properties"] + ul'
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
      if self[x] != nil && self[x] != y
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

    ms = @classes.map { |clsn, kvs| kvs.map { |mem, kvs2| Hash === kvs2 && kvs2.has_key?("m_doc") ? "#{clsn}.#{mem}" : nil } }.flatten.compact
    ps = @classes.map { |clsn, kvs| kvs.map { |mem, kvs2| Hash === kvs2 && kvs2.has_key?("p_doc") ? "#{clsn}.#{mem}" : nil } }.flatten.compact

    all_data = {
      "TODO" => "find  ': null' and '#TODO FIXME' and '__UN'   check something problem",
      N_SUMM => {
        downloads: @links_all.size,
        parsed: @classes.size + ms.size + ps.size,
        parse_skips: @url_info_kvs.select { |url, kvs| kvs.has_key?("UN_PROCESS") }.size,
        N_co => @collection.size,
        N_e => @enumeration.size,
        N_c => @classes.size,
        N_m => ms.size,
        N_p => ps.size,
        enumeration_table_not_found: @enumeration_table_not_found,
      },
      N_co => @collection,
      N_e => @enumeration,
      N_c => @classes,
      "#{N_c}_#{N_m}" => ms,
      "#{N_c}_#{N_p}" => ps,
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
    page_title = hh.css ::HTMLCSS::H1_TITLE
    if page_title.size < 1
      @url_info_kvs[url]["UN_PROCESS"] = "todo NOT GET H1_TITLE"
      return
    end
    app_name, objtype, obj_name = add_url_info(url, page_title[0].text)
    # TODO FIXME obj_name may has '.'

    if app_name != ::OLE_APPNAME::EXCEL
      @url_info_kvs[url]["UN_PROCESS"] = "todo NOT PROCESS APP"
      return
    end

    iffo = { N_META => { url: url, fn: File.basename(url_fn(url)), size: txt.size } }

    case objtype
    when ::OLE_TYPE::PROPERTY
      parse_property_html(url, hh, obj_name, iffo)
    when ::OLE_TYPE::METHOD
      parse_method_html(url, hh, obj_name, iffo)
    when ::OLE_TYPE::OBJECT
      es, ms, ps = parse_object_html(url, hh, obj_name, iffo)
      # TODO not need
      # es.map { |t,href| @links_all[href] = ::OLE_TYPE::EVENT }
      ps.map { |t, href| @links_all[href] = ::OLE_TYPE::PROPERTY }
      ms.map { |t, href| @links_all[href] = ::OLE_TYPE::METHOD }
    when ::OLE_TYPE::COLLECTION
      parse_collection_html(url, hh, app_name, objtype, obj_name, iffo)
    when ::OLE_TYPE::ENUMERATION
      parse_enumeration_html(url, hh, obj_name, iffo)
    else
      @url_info_kvs[url]["UN_PROCESS"] = "todo EXCEL OBJ OTHER"
    end
    @parsed_html += 1
  end

  protected

  def assert(x, msg = "")
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

  def foreach_a_add_to_links(url, *vsss)
    vsss.each do |vss|
      vss.each do |ret|
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

  def parse_collection_html(url, hh, app_name, objtype, obj_name, iffo)
    if app_name == ::OLE_APPNAME::EXCEL && objtype == ::OLE_TYPE::COLLECTION
      obj_name = "VBA_" + ::OLE_TYPE::COLLECTION
    end
    assert(@collection[obj_name] == nil)
    @collection[obj_name] = iffo.update({ N_co_doc => nil })

    docss = hh.css ::HTMLCSS::ID_DOC_p_pres
    if docss.size > 0
      iffo[N_co_doc] = docss.map { |e| e.text.strip }.join("\n\n")
    end
    foreach_a_add_to_links(url, docss)
  end

  def parse_enumeration_html(url, hh, obj_name, iffo)
    assert(@enumeration[obj_name] == nil)
    @enumeration[obj_name] = iffo.update({ N_e_doc => nil })

    docss1 = hh.css ::HTMLCSS::ID_DOC_p
    if docss1.size > 0
      iffo[N_e_doc] = docss1[0].text.strip
    end

    docss2 = hh.css ::HTMLCSS::ID_DOC_table
    if docss2.size > 0
      ts = hh.css "table"
      tab1 = docss2[0]

      th = tab1.css("th").map(&:text)
      if th[0..2] != %w{Name Value Description}
        iffo[N_e_table_format_incorrect] = true
      end
      rows = tab1.css("tr").map { |tr| tr.css("td").map(&:text) } - [[]]
      rows = rows.map do |row|
        if row[1].include?(" ")
          row[1] = row[1].sub(" ", " #TODO FIXME") # Most enumeration is Long
        end
        row
      end
      iffo[N_e_table_head] = th
      iffo[N_e_table_rows] = rows
      vs1 = rows.map { |row| row[1] }
      vs2 = vs1.uniq
      if vs1 == vs2
        iffo[N_e_unique] = true
      end
    else
      iffo[N_e_table_head] = []
      iffo[N_e_table_rows] = []
      @enumeration_table_not_found += 1
    end

    remarkss = hh.css ::HTMLCSS::ID_REMARKS_ps
    if remarkss.size > 0
      iffo[N_e_remarks_doc] = remarkss.map { |e| e.text.strip }.join("\n\n")
    end
    foreach_a_add_to_links(url, docss1, docss2, remarkss)
  end

  def parse_property_html(url, hh, prop_name, iffo)
    nns = prop_name.split(".")
    assert(nns.size == 2)
    clsn, mem = nns
    @classes[clsn] ||= MyHash.new
    assert(@classes[clsn][mem] == nil)
    @classes[clsn][mem] = iffo.update({ N_p_doc => nil, N_p_type => nil })

    returnss = hh.css ::HTMLCSS::ID_RETURN_VALUE_p
    if returnss.size > 0
      ty = returnss[0].text.strip
      iffo[N_p_return_doc] = ty
      ty = find_property_type_name(returnss, url)
      if ty
        iffo[N_p_type] = ty
      else
        iffo[N_p_type] = "__UN_PARSED_RETURN_VALUE__"
      end
    end

    propss = hh.css ::HTMLCSS::ID_PROPERTY_VALUE_p
    if propss.size > 0
      ty = propss[0].text.strip
      iffo[N_p_property_value_doc] = ty
      ty = find_property_type_name(propss, url)
      if ty
        iffo[N_p_type] = ty
      else
        iffo[N_p_type] = "__UN_PARSED_PROPERTY_VALUE__"
      end
    end

    docss1 = hh.css ::HTMLCSS::ID_DOC_p
    if docss1.size > 0
      iffo[N_p_doc] = docss1[0].text.strip
      ty = find_property_type_name(docss1, url)
      iffo[N_p_type] = ty if !iffo[N_p_type] && ty
    else
      docss1 = hh.css ::HTMLCSS::ID_DOC_div
      if docss1.size > 0
        iffo[N_p_doc] = docss1[0].text.strip
        ty = find_property_type_name(docss1, url)
        iffo[N_p_type] = ty if !iffo[N_p_type] && ty
      end
    end

    syntaxss = hh.css ::HTMLCSS::ID_SYNTAX_ps
    if syntaxss.size > 0
      iffo[N_p_syntax_doc] = syntaxss.map { |e| e.text.strip }.join("\n\n")
      if !iffo[N_p_type]
        ty = find_property_type_name(syntaxss, url)
        iffo[N_p_type] = ty if ty
      end
    end

    remarkss = hh.css ::HTMLCSS::ID_REMARKS_ps
    if remarkss.size > 0
      iffo[N_p_remarks_doc] = remarkss.map { |e| e.text.strip }.join("\n\n")
    end

    exampless = hh.css ::HTMLCSS::ID_EXAMPLE_p_pres
    if exampless.size > 0
      iffo[N_p_example_doc] = exampless.map { |e| e.text.strip }.join("\n\n")
    end

    foreach_a_add_to_links(url, returnss, propss, docss1, syntaxss, remarkss, exampless)
    if iffo[N_p_type] && iffo[N_p_type].include?("-")
      iffo[N_p_type].gsub!("-", "")
    end
    iffo[N_p_type] = "__UNKNOWN_TYPE__" unless iffo[N_p_type]
  end

  def proc_aas(res)
    tys = res.css("strong,a")
      .map { |e| e.text }
      .reject { |e| e.include?(" ") }
      .map { |e| VBAtype2pytype.has_key?(e) ? VBAtype2pytype[e] : e }
      .uniq.join(" | ")
    tys = "Any | None" if tys == "None"
    tys
  end

  def find_property_type_name(vss, url)
    # TODO FIXME

    vss.map do |res|
      ty = res.text.strip

      if %r{Note.*is deprecated and is not intended}im =~ ty
        return "__DEPRECATED_WARNNING__"
      end

      if VBAtype2pytype.has_key?(ty)
        return VBAtype2pytype[ty]
      elsif ty !~ /\s|\./ # TODO
        return ty
      end

      if %r{Can be.*Read/write Long}i =~ ty
        tys = proc_aas(res)
        return tys if tys != ""
      end
      if %r{Read/write Long}i =~ ty
        tys = proc_aas(res)
        return tys if tys != ""
      end
      if %r{Read-only\.}i =~ ty
        tys = proc_aas(res)
        return tys if tys != ""
      end
      if %r{(?:Read-only|Read/write) (\w+)}i =~ ty
        clsn = $1
        if VBAtype2pytype.has_key?(clsn)
          return VBAtype2pytype[clsn]
        elsif /\s/ !~ clsn
          return clsn
        end
      end

      if %r{^\w+ object}i =~ ty
        return VBAtype2pytype["object"]
      end
      if %r{Returns an (\w+).*that represents}i =~ ty || %r{Returns a (\w+) value}i =~ ty
        e = $1
        e = [VBAtype2pytype.has_key?(e) ? VBAtype2pytype[e] : e]
        if %r{Returns Nothing} =~ ty
          e << "None"
        end
        return e.join(" | ")
      end
      if %r{Returns an object that represents.*Returns Nothing}i =~ ty
        tys = [VBAtype2pytype["object"], "None"]
        return tys.join(" | ")
      end
      if %r{Returns an? .*? object that represents the}i =~ ty
        return VBAtype2pytype["object"]
      end
      if %r{specified object.*Read.*\.}i =~ ty
        return VBAtype2pytype["object"]
      end
      if %r{Returns an? single object}i =~ ty
        return VBAtype2pytype["object"]
      end
      if %r{Read-only\.}i =~ ty
        return VBAtype2pytype["object"]
      end
      if %r{Returns the array}i =~ ty
        return VBAtype2pytype["array"]
      end
      if %r{Returns or sets a (\w+) value}i =~ ty
        e = $1
        return VBAtype2pytype.has_key?(e) ? VBAtype2pytype[e] : e
      end
    end

    %w{a strong}.each do |xxx|
      tys = vss.map { |e| e.css xxx }
        .flatten
        .map { |e| e.text }
        .reject { |e| e.include?(" ") }
        .map { |e| VBAtype2pytype.has_key?(e) ? VBAtype2pytype[e] : e }
        .uniq

      if tys.size == 0
      elsif tys.size == 1
        return tys[0]
      else
        return tys.join(" | ")
      end
    end

    nil
  end

  def find_method_type_name(vss, url)
    # TODO FIXME
    vss.map do |res|
      ty = res.text.strip
      if VBAtype2pytype.has_key?(ty)
        return VBAtype2pytype[ty]
      elsif ty !~ /\s|\./ # TODO
        return ty
      end

      tys = proc_aas(res)
      return tys if tys != ""

      if %r{Creates an? new }i =~ ty
        return VBAtype2pytype["object"]
      end
      if %r{An? Object value that represents}i =~ ty
        return VBAtype2pytype["object"]
      end
      if %r{Returns an? single object}i =~ ty
        return VBAtype2pytype["object"]
      end
    end

    %w{a strong}.each do |xxx|
      tys = vss.map { |e| e.css xxx }
        .flatten
        .map { |e| e.text }
        .reject { |e| e.include?(" ") }
        .map { |e| VBAtype2pytype.has_key?(e) ? VBAtype2pytype[e] : e }
        .uniq

      if tys.size == 0
      elsif tys.size == 1
        return tys[0]
      else
        return tys.join(" | ")
      end
    end

    nil
  end

  def parse_method_html(url, hh, method_name, iffo)
    nns = method_name.split(".")
    assert(nns.size == 2)
    clsn, mem = nns
    @classes[clsn] ||= MyHash.new
    assert(@classes[clsn][mem] == nil)
    @classes[clsn][mem] = iffo.update({ N_m_doc => nil, N_m_return => nil })

    returnss = hh.css ::HTMLCSS::ID_RETURN_VALUE_p
    if returnss.size > 0
      ty = find_method_type_name(returnss, url)
      if ty
        iffo[N_m_return] = ty
      else
        iffo[N_m_return] = "__UN_PARSED_TYPE__" # un parsed string
      end
      iffo[N_m_return_doc] = returnss.text.strip
    else
      iffo[N_m_return] = "None"
    end

    docss = hh.css ::HTMLCSS::ID_DOC_p
    if docss.size > 0
      iffo[N_m_doc] = docss[0].text.strip
      unless iffo[N_m_return]
        ty = find_method_type_name(docss, url)
        iffo[N_m_return] = ty if ty
      end
    end

    parameterss = hh.css ::HTMLCSS::ID_PARAMETERS_table
    if parameterss.size > 0
      iffo[N_m_parameters_doc] = parameterss[0].text.gsub(/((\r)?\n){3,}/, "\n\n").gsub(/(?<!\n)\n(?!\n)/, " ").strip
    end

    remarkss = hh.css ::HTMLCSS::ID_REMARKS_ps
    if remarkss.size > 0
      iffo[N_m_remarks_doc] = remarkss.map { |e| e.text.strip }.join("\n\n")
    end

    exampless = hh.css ::HTMLCSS::ID_EXAMPLE_p_pres
    if exampless.size > 0
      iffo[N_m_example_doc] = exampless.map { |e| e.text.strip }.join("\n\n")
    end

    if iffo[N_m_return] && iffo[N_m_return].include?("-")
      iffo[N_m_return].gsub!("-", "")
    end
    iffo.delete(N_m_return) if iffo[N_m_return] == nil
  end

  def parse_object_html(url, hh, cls_name, iffo)
    @classes[cls_name] ||= MyHash.new
    # @classes[cls_name] = iffo.update({ N_c_doc => nil })
    assert(!@classes[cls_name].has_key?(N_c_doc))
    @classes[cls_name].update(iffo.update({ N_c_doc => nil }))
    iffo = @classes[cls_name]

    docss = hh.css ::HTMLCSS::ID_DOC_p
    if docss.size > 0
      iffo[N_c_doc] = docss[0].text.strip
    end
    remarkss = hh.css ::HTMLCSS::ID_REMARKS_ps
    if remarkss.size > 0
      iffo[N_c_remarks_doc] = remarkss.map { |e| e.text.strip }.join("\n\n")
    end

    exampless = hh.css ::HTMLCSS::ID_EXAMPLE_p_pres
    if exampless.size > 0
      iffo[N_c_example_doc] = exampless.map { |e| e.text.strip }.join("\n\n")
    end
    foreach_a_add_to_links(url, docss, remarkss, exampless, docss, remarkss)

    eventss = hh.css ::HTMLCSS::ID_EVENTS_ul
    assert(eventss.size <= 1, %{eventss ul > 1, #{eventss.size}  #{url}})
    methodss = hh.css ::HTMLCSS::ID_METHODS_ul
    assert(methodss.size <= 1, %{methodss ul > 1, #{methodss.size}  #{url}})
    propertiess = hh.css ::HTMLCSS::ID_PROPERTIES_ul
    assert(propertiess.size <= 1, %{propertiess ul > 1, #{propertiess.size}  #{url}})

    baseurl = url_base(url)
    eventss = if eventss.size > 0
        a2arr(baseurl, eventss[0].css("a"))
      else
        []
      end
    methodss = if methodss.size > 0
        a2arr(baseurl, methodss[0].css("a"))
      else
        []
      end
    propertiess = if propertiess.size > 0
        a2arr(baseurl, propertiess[0].css("a"))
      else
        []
      end
    STDERR.puts %{ HTML #{txt.size} bytes. SUMMARY: events #{eventss.size},  methods #{methodss.size}, properties #{propertiess.size}} if $DEBUG

    [eventss, methodss, propertiess]
  end
end

if __FILE__ == $0
  Dir.mkdir(CACHEDIR) unless File.exist?(CACHEDIR)

  begin_url = "https://learn.microsoft.com/en-us/office/vba/api/excel.application(object)"

  dd = DownAPI.new(begin_url)
  dd.download_all()
end
