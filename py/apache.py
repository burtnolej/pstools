from apachelogs import LogParser
import dns.resolver, dns.reversename
import gzip
from shutil import copyfileobj



ignore={"44.197.209.225":{"host":"ec2-44-197-209-225.compute-1.amazonaws.com","count":0}}
parser = LogParser("%h %l %u %t \"%r\" %>s %b \"%{Referer}i\" \"%{User-Agent}i\"")

with open("./logs/access.log","r+") as f:
    for line in f:
        entry = parser.parse(line)
        if entry.remote_host not in ignore.keys():
            try:
                addrs = dns.reversename.from_address(entry.remote_host)
                source = str(dns.resolver.resolve(addrs,"PTR")[0])
                print(source,entry.request_line,entry.request_time)   
            except:
                print("error",entry.remote_host)
        else:
            ignore[entry.remote_host]["count"]=ignore[entry.remote_host]["count"]+1
            
print(ignore)