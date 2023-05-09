kvs2 = {
    "aa":{
        "title":"asdf",
        "type":"af"
    },
    "ab":{
        "title":"asdf",
        "type":"af2"
    }
}

kvs2 = [
  [
    "https://learn.microsoft.com/en-us/office/vba/api/excel.addin",
    {
      "title": "AddIn object (Excel)",
      "type": "object (Excel)",
      "htmlf": "exceldoc/excel.addin.html"
    }
  ],
  [
    "https://learn.microsoft.com/en-us/office/vba/api/excel.addin.application",
    {
      "title": "AddIn.Application property (Excel)",
      "type": "property (Excel)",
      "htmlf": "exceldoc/excel.addin.application.html"
    }
  ]].to_h

type_kvs =  kvs2.group_by{|url,kvs| kvs['type']}

all_data = {
url_info_kvs_summ: %{#{kvs2.size} [ #{type_kvs.map { |ty,v| ty.to_s+':'+v.size.to_s }.join(', ') } ] },
kvs: type_kvs,
kvs2: kvs2,
}

require'json'

puts JSON.pretty_generate(all_data)
