require'set'
CACHEDIR='exceldoc'
require'nokogiri'


require_relative '../downapi'


class DD2 < DownAPI


  def nnn
    url='https://learn.microsoft.com/en-us/office/vba/api/excel.range.numberformatlocal'

    aas=%q{<a href="https://support.office.com/article/number-format-codes-5026bbd6-04bc-48cd-bf33-80f18b4eae68" data-linktype="external">Number format codes (Microsoft Support)</a>}

    hh = Nokogiri::HTML aas
    a = hh.css( 'a')[0]
    # @links_skips = Set.new


    aurl = a['href']
    x1=url_base(url)
    p [url, x1, aurl  , a.to_s]
    url2 = baseurl_to_fullurl(x1, aurl)
    p ['url2', url2 ]

    if url2
      @links_all[url2] = HtmlType::OBJECT
    else
      STDERR.puts %{WHY nil　 url:#{url}  url2:#{url2}    #{a.to_s}}
    end
    p a.text.strip
  end

end

DD2.new.nnn
