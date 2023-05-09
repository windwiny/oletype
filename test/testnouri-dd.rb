
def url_nor(u1)
  #u1='https://learn.microsoft.com/en-us/office/vba/api/../excel/concepts/workbooks-and-worksheets/../../../api/excel.xmlnamespace.prefix'
  #u2='https://learn.microsoft.com/en-us/office/vba/api/excel.xmlnamespace.prefix'
  vs=u1.split('/')
  2.upto(vs.size).each do |i|
    if vs[i]=='..'
      vs[i]=nil
      j=i-1
      while vs[j] == nil
        j=j-1
      end
      vs[j]=nil
    end
  end
  u2=vs.compact.join('/')
  u2
end

puts url_nor(u1)
