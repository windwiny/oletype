require'json'
require'set'
require'stringio'
require'net/http'
require'nokogiri'


# TODO
HTMLLINKS_FN = 'htmlpagelink2fn.txt'
API_FN = 'excel.info.json'
DOCDIR = 'exceldoc'

CACHEDIR = File.join(__dir__, DOCDIR)
Dir.mkdir(CACHEDIR) unless File.exist?(CACHEDIR)

# TODO FIXME      vba's type to python type
VBAtype2pytype = {
  'Boolean'=>'bool',
  'Byte'=>'VBA_Byte',
  'Currency'=>'VBA_Currency',
  'Date'=>'datetime.datetime',
  'Decimal'=>'VBA_Decimal',
  'Double'=>'float',
  'Integer'=>'int',
  'Long'=>'int',
  'LongLong'=>'int',
  'LongPtr'=>'int',
  'Object'=>'VBA_Object',
  'Single'=>'float',
  'String'=>'str',
  'Variant'=>'VBA_Variant',

  'Nothing'=>'None',
  'VOID'=>'None',
  'Void'=>'None',
  'OBJECT'=>'VBA_OBJECT',
  'object'=>'VBA_object',
  'True'=>'bool',
  'False'=>'bool',
}

def manual_find_type_from_string(ss)
  ## TODO FIXME
  ## see output xxx.info.json files, some property/method type not found

  ss
end

class MyHash < Hash
  def []=(x, y)
    if self.has_key? x
      raise "ERRR"
    end
    super
  end
end

class HtmlType
  OBJECT = 'OBJECT'
  EVENT = 'EVENT'
  METHOD = 'METHOD'
  PROPERTY = 'PROPERTY'
end

class DownAPI

  def initialize start_url
    @links_all = Hash.new()
    @links_all[start_url] = HtmlType::OBJECT

    @undown = []
    @downloaded = Set.new()

    @classes = MyHash.new
    @properties = MyHash.new
    @methods = MyHash.new
  end

  def download_all()
    @undown = @links_all.to_a
    while true
      @undown.each do |url, objtype|
        if @downloaded.include?(url)
          next
        end
        @downloaded.add(url)
        process_page(url, objtype)
      end

      if @downloaded.to_a == @links_all.keys()   # all links downed
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
    STDERR.puts(%{OUTPUT files})
    all_data = {
      summary: {
        classes: @classes.size,
        methods: @methods.size,
        properties: @properties.size,
      },
      classes: @classes,
      methods: @methods,
      properties: @properties,
    }

    File.write(API_FN,  JSON.pretty_generate(all_data) )

    type_kvs = @links_all.group_by{|k,v| v}
    ss = [%{## SUMM: #{@links_all.size},   #{type_kvs.map {|k,v| %{#{k}:#{v.size}}}.join(', ')}} ]
    type_kvs.each do |k,v|
      ss << %{# #{k} #{v.size}}
      ss << v.map do |x|
        url = x[0]
        fn = url_fn(url)
        fn = %{#{DOCDIR}/#{File.basename(fn)}}
        fninfo = if File.exist?(fn)
          fn
        else
          %{NOT EXIST #{fn}}
        end
        %{#{url} -> #{fninfo}}
      end.sort
    end
    sss = ss.flatten.join("\n")
    if !(File.exist?(HTMLLINKS_FN) && File.read(HTMLLINKS_FN) == sss)
      File.write(HTMLLINKS_FN, sss)
    end
  end


protected

  def process_page(url, objtype)
    txt = download_page(url)
    case objtype
    when HtmlType::EVENT
      parse_event_html(url, txt)
    when HtmlType::PROPERTY
      parse_property_html(url, txt)
    when HtmlType::METHOD
      parse_method_html(url, txt)
    else # HtmlType::OBJECT
      es, ms, ps= parse_object_html(url, txt)
      # TODO not need
      # es.map { |t,href| @links_all[href] = HtmlType::EVENT }
      ps.map { |t,href| @links_all[href] = HtmlType::PROPERTY }
      ms.map { |t,href| @links_all[href] = HtmlType::METHOD }
    end
  end



protected
  def assert x, msg
    if !x
      raise %{#{msg} Failed}
    end
  end

  def url_split(url)
    url = url.to_s
    ix = url.rindex('/')
    if ix
      base = url[..ix]
      fn = url[ix+1..]
    else
      base = ''
      fn = url
    end
    fn = fn.gsub(/[\\\/\:\*\?\"\<\>\|]/, '_') # illegal filename character on win \/:*?"<>|
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
    url = if url.include?('/') && url[0] != '.'
      url
    else
      baseurl + url
    end
    url
  end

  def a2arr(baseurl, ess)
    res = ess.map do |e|
      href = baseurl_to_fullurl(baseurl, e['href'])
      [e.text, href]
    end
    res
  end

  def save_html(url, txt)
    fn = url_fn(url)
    if !txt || txt.strip == ''
      STDERR.puts %{SKIP EMPTY #{url}} if $DEBUG
      return
    end
    STDERR.puts %{ SAVE #{File.basename(fn)},  #{txt.size}} if $DEBUG
    File.write(fn, txt)
  end

  def download_page(url)
    fn = url_fn(url)
    if File.exist?(fn)
      STDERR.puts %{SKIP #{url} } if $DEBUG
      txt = File.read(fn)
      return txt
    end

    url = URI(url)
    STDERR.puts %{DOWNLOADING #{url} }
    begin
      r = Net::HTTP.get_response( url)
      txt = r.body
      save_html(url, txt)
      return txt
    rescue Exception => e
      STDERR.puts %{Net::HTTP ERROR #{e} on #{url}}
    end
    ''
  end

  def parse_event_html(url, txt)
    h = Nokogiri::HTML txt
    raise 'No implement yet'  # TODO FIXME
  end

  def find_type_name(returns, url)
    aas = returns[0].css('a')
    if aas.size > 0   # return object, find sit
      assert(aas.size > 0, "#{url_fn(url)} return what?")

      a = aas[0]
      returnobj = a.text
      aurl = a['href']
      @links_all[baseurl_to_fullurl(url_base(url), aurl)] = HtmlType::OBJECT
    else
      r2 = returns[0].css('strong')
      if r2.size > 0
        returnobj = r2[0].text
      else
        r3 = returns[0].css('b')
        if r3.size > 0
          returnobj = r3[0].text
        else
          returnobj = returns[0].text
        end
      end
      # Manual check
      if returnobj =~ /True.*False/
        returnobj = 'bool'
      elsif VBAtype2pytype.has_key?(returnobj)
        returnobj = VBAtype2pytype[returnobj]
      else
        returnobj = manual_find_type_from_string(returnobj)
      end
    end
    returnobj.strip
  end

  def parse_property_html(url, txt)
    h = Nokogiri::HTML txt
    cls_member = h.css 'div > h1'
    assert(cls_member.size == 1, %{not get Cls.Property title  #{url}})
    clsn, member = cls_member[0].text.split()[0].split('.')

    iffo = {}
    returns = h.css 'nav[id="center-doc-outline"] + p'
    if returns.size > 0
      iffo['type'] = find_type_name(returns, url)
      iffo['info'] = returns[0].text.strip
    else
      returns = h.css 'nav[id="center-doc-outline"] + div'
      if returns.size > 0
        iffo['type'] = find_type_name(returns, url)
        iffo['info'] = returns[0].text.strip
      else
        iffo['type'] = nil
        iffo['info'] = nil
        iffo['ERROR'] = %{not found info on '#{File.basename url_fn url}' }
      end
    end
    @properties[%{#{clsn}.#{member}}] = iffo
  end

  def parse_method_html(url, txt)
    h = Nokogiri::HTML txt
    cls_member = h.css 'div > h1'
    assert(cls_member.size == 1, %{not get Cls.Method title  #{url}})
    clsn, member = cls_member[0].text.split()[0].split('.')

    iffo = {}
    returns = h.css 'nav[id="center-doc-outline"] + p'
    if returns.size > 0
      iffo['method'] = returns[0].text.strip
    else
      iffo['method'] = %{not found comment on '#{File.basename url_fn url}' }
    end

    returns = h.css 'div > h2[id="return-value"] + p'
    if returns.size > 0
      iffo['return'] = find_type_name(returns, url)
    else
      iffo['return'] = nil
    end
    @methods[%{#{clsn}.#{member}}] = iffo
  end

  def parse_object_html(url, txt)
    h = Nokogiri::HTML txt

    cls_member = h.css 'div > h1'
    if cls_member.size < 1
      STDERR.puts %{not title , skip #{url}}
      return [[],[],[]]
    end
    clsn = cls_member[0].text.split()[0].split('.')[-1]

    iffo = {}
    returns = h.css 'nav[id="center-doc-outline"] + p'
    if returns.size > 0
      iffo['class'] = returns[0].text.strip
    end
    @classes[clsn] = iffo

    events = h.css 'div > h2[id="events"] + ul'
    assert(events.size <= 1, %{events ul > 1, #{events.size}  #{url}})
    methods = h.css 'div > h2[id="methods"] + ul'
    assert(methods.size <= 1, %{methods ul > 1, #{methods.size}  #{url}})
    properties = h.css 'div > h2[id="properties"] + ul'
    assert(properties.size <= 1, %{properties ul > 1, #{properties.size}  #{url}})

    baseurl = url_base(url)
    events = if events.size > 0
      a2arr(baseurl, events[0].css('a'))
    else
      []
    end
    methods = if methods.size > 0
      a2arr(baseurl, methods[0].css('a'))
    else
      []
    end
    properties = if properties.size > 0
      a2arr(baseurl, properties[0].css('a'))
    else
      []
    end
    STDERR.puts %{ HTML #{txt.size} bytes. SUMMARY: events #{events.size},  methods #{methods.size}, properties #{properties.size}} if $DEBUG


    [events, methods, properties]
  end
end

begin_url = 'https://learn.microsoft.com/en-us/office/vba/api/excel.application(object)'

dd = DownAPI.new(begin_url)
dd.download_all()

dd.finish()

