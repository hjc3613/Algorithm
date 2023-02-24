egrep '[(（]九[）)]\s*其他重大风险' 688*.csv | awk -F ':' '{print $1}' > has_key.txt
comm -23 <(sort all_key.txt) <(sort has_key.txt)