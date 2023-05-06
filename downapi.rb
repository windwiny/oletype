require'json'
require'set'
require'stringio'
require'net/http'
require'nokogiri'

CACHEDIR=File.join(__dir__, 'exceldoc')
Dir.mkdir(CACHEDIR) unless File.exist?(CACHEDIR)

class HtmlType
  OBJECT = 'OBJECT'
  EVENT = 'EVENT'
  METHOD = 'METHOD'
  PARAMETER = 'PARAMETER'
end

api_comment_fn = 'excel.apicomment.json'
$api_comment_kvs = {}

# TODO FIXME      vba's type to python type
VBAtype2pytype = {
  'Boolean'=>'bool',
  'Byte'=>'str',
  'Currency'=>'float',
  'Date'=>'datetime.datetime',
  'Decimal'=>'float',
  'Double'=>'float',
  'Integer'=>'int',
  'Long'=>'int',
  'LongLong'=>'int',
  'LongPtr'=>'int',
  'Object'=>'object',
  'Single'=>'float',
  'String'=>'str',
  'Variant'=>'list',

  'Nothing'=>'None',
  'VOID'=>'None',
  'Void'=>'None',
  'OBJECT'=>'object',
  'object'=>'object',
}

class DownAPI

  def initialize start_url
    @links_all = Hash.new()
    @links_all[start_url] = HtmlType::OBJECT

    @undown = []
    @downloaded = Set.new()
    @parameter_type = StringIO.new
    @method_return = StringIO.new
  end

  def download_all()
    @undown =  @links_all.to_a
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
    puts
    puts "ALL Class.Method => Return:"
    puts @method_return.string
    puts @method_return.string.lines.size
    puts
    puts "ALL Class.Parameter => Type:"
    puts @parameter_type.string
    puts @parameter_type.string.lines.size
  end


protected

  def process_page(url, objtype)
    txt = download_page(url)
    case objtype
    when HtmlType::EVENT
      parse_event_html(url, txt)
    when HtmlType::PARAMETER
      parse_parameter_html(url, txt)
    when HtmlType::METHOD
      parse_method_html(url, txt)
    else # HtmlType::OBJECT
      es, ms, ps= parse_object_html(url, txt)
      # TODO not need
      # es.map { |t,href| @links_all[href] = HtmlType::EVENT }
      ps.map { |t,href| @links_all[href] = HtmlType::PARAMETER }
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
    fn = fn.gsub(/\W/, '_')
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

  def parse_parameter_html(url, txt)
    h = Nokogiri::HTML txt
    cls_member = h.css 'div > h1'
    assert(cls_member.size == 1, 'not get Cls.Method title')
    clsn, member = cls_member[0].text.split()[0].split('.')

    returns = h.css 'nav[id="center-doc-outline"] + p'
    assert(returns.size <= 1, "get parameter type ,size!=1 #{returns.size}")
    if returns.size > 0
      aas = returns[0].css('a')
      if aas.size > 0   # return object, find sit
        assert(aas.size > 0, "#{url_fn(url)} return what?")

        a = aas[0]
        returnobj = a.text
        aurl = a['href']
        @links_all[baseurl_to_fullurl(url_base(url), aurl)] = HtmlType::OBJECT
      else
        returnobj = returns[0].text
        # Manual check
        if returnobj =~ /True.*False/
          returnobj = 'bool'
        elsif VBAtype2pytype.has_key?(returnobj)
          returnobj = VBAtype2pytype[returnobj]
        end
      end
      $api_comment_kvs[%{#{clsn}_#{member}}] = returns[0].text
    else
      returnobj = ''
    end
    @parameter_type.puts %{  #{clsn}.#{member} -> #{returnobj}}
  end

  def parse_method_html(url, txt)
    h = Nokogiri::HTML txt
    cls_member = h.css 'div > h1'
    assert(cls_member.size == 1, 'not get Cls.Method title')
    clsn, member = cls_member[0].text.split()[0].split('.')

    returns = h.css 'nav[id="center-doc-outline"] + p'
    if returns.size > 0
      $api_comment_kvs[%{#{clsn}_#{member}}] = returns[0].text
    end

    returns = h.css 'div > h2[id="return-value"] + p'
    if returns.size > 0

      aas = returns[0].css('a')
      if aas.size > 0   # return object, find sit
        assert(aas.size > 0, "#{url_fn(url)} return what?")

        a = aas[0]
        returnobj = a.text
        aurl = a['href']
        @links_all[baseurl_to_fullurl(url_base(url), aurl)] = HtmlType::OBJECT
      else
        returnobj = returns[0].text
        # Manual check
        if returnobj =~ /True.*False/
          returnobj = 'bool'
        elsif VBAtype2pytype.has_key?(returnobj)
          returnobj = VBAtype2pytype[returnobj]
        end
      end
    else
      returnobj = ''
    end
    @method_return.puts %{  #{clsn}.#{member} -> #{returnobj}}
  end

  def parse_object_html(url, txt)
    h = Nokogiri::HTML txt

    events = h.css 'div > h2[id="events"] + ul'
    assert events.size <= 1, "events ul > 1, #{events.size}"
    methods = h.css 'div > h2[id="methods"] + ul'
    assert methods.size <= 1, "methods ul > 1, #{methods.size}"
    properties = h.css 'div > h2[id="properties"] + ul'
    assert properties.size <= 1, "properties ul > 1, #{properties.size}"

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

File.write(api_comment_fn,  JSON.pretty_generate($api_comment_kvs) )
