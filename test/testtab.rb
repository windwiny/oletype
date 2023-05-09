hs=%{
    <table>
<thead>
<tr>
<th style="text-align: left;">Name</th>
<th style="text-align: left;">Value</th>
<th style="text-align: left;">Description</th>
</tr>
</thead>
<tbody>
<tr>
<td style="text-align: left;"><strong>xlEqualAllocation</strong></td>
<td style="text-align: left;">1</td>
<td style="text-align: left;">Use equal allocation.</td>
</tr>
<tr>
<td style="text-align: left;"><strong>xlWeightedAllocation</strong></td>
<td style="text-align: left;">2</td>
<td style="text-align: left;">Use weighted allocation.</td>
</tr>
</tbody>
</table>
}

require 'nokogiri'

hh = Nokogiri::HTML hs

ts = hh.css 'table'
tab1 = ts[0]

th = tab1.css('th').map( &:text)
puts th.join("\t")
p
rows = tab1.css('tr').map { |tr|  tr.css('td').map( &:text)  }
rows.each do |row|
  puts row.join("\t")
end
