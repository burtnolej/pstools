import sys
import datetime
import ast
import time

all_categories=["Job Type","Sub Department","Seniority","Department"]

for cfg in sys.argv[1:]:
        (k,v) = cfg.split("=")
        if v in ["True","true"]: v =True
        if v in ["False","false"]: v =False
        locals()[k] = v


def _get_client_meta(client_name,clients):

        client_type = "UNKNOWN"
        full_client_name = "UNKNOWN"

        if client_name in clients.keys():
                client_type = clients[client_name]["Company Type"]
                
                if "Notes" in clients[client_name].keys():
                        try:
                                full_client_name = clients[client_name]["Notes"].split(";")[0].split("=")[1]
                        except:
                                pass

        return(client_type,full_client_name)

def _get_person_meta(results,_id):
        result=[results[_id]["firstname"],
                results[_id]["lastname"], \
                results[_id]["email"], \
                _id, \
                results[_id]["phone"], \
                results[_id]["organization"], \
                results[_id]["jobtitle"]]
        
        if "debug" in results[_id]['results'].keys():
                result.append("^".join(results[_id]['results']["debug"]))

        for _category in all_categories:
                result.append(results[_id]["results"][_category])

        return result


def print_results(persons,clients):
        fh = open(outputfile,"w")
        _header=["firstName","lastName","emailAddress","personId","phone","organisation","jobTitle","debugSeniority","debugDepartment","debugSubDepartment","debugJobType","Job Type","Sub Department","Seniority","Department"]
        
        fh.write("^".join(_header)+"\n")
        for id in persons.keys():
                #(client_type,full_client_name) = _get_client_meta(results[id]["organization"],clients)
                _output = _get_person_meta(persons,id)
                #_output.append(client_type)
                #_output.append(full_client_name)
                fh.write("^".join(_output)+"\n")
        fh.close

def get_rulesets(rulesfile, delimiter="^"):
        rulesets={}
        fh = open(rulesfile,"r+")
        for line in fh:
                line=line.replace("\"","")
                line=line.replace("\r","")
                line=line.replace("\n","")
                _line =line.split(delimiter)

                if _line[0]!="":
                        if _line[0] in rulesets.keys():
                                rulesets[_line[0]].append(_line)
                        else:	
                                rulesets[_line[0]] = [line.split(delimiter)]
        fh.close()
        return rulesets

#name^phoneNumbers^team^owner^emailAddresses^id^createdAt^updatedAt^Company Type^Company Size^Head Region

def get_clients(clientsfile):
        clients={}
        fh = open(clientsfile,"r+")
        linecount=0
        for line in fh:
                if linecount!=0:
                        _client = line.split("^")
                        client = {"id":_client[5]}
                        client["Company Type"] = _client[8]
                        client["Company Size"] = _client[9]
                        client["Company Region"] = _client[10]
                        client["Notes"] = _client[11]
                        _name = _client[0]

                        clients[_name] = client
                else:
                        linecount=linecount+1
        fh.close()
        return clients

#emailAddresses^Contact Owner^firstName^id^jobTitle^lastName^organisation^owner^phoneNumbers^team^title^lastContactedAt^createdAt^updatedAt^Job Type^Department^Sub Department^Seniority^LinkedInURL^Notes

def get_persons(personsfile):
        persons=[]
        fh = open(personsfile,"r+")
        linecount=0
        for line in fh:
                if linecount!=0:
                        _person = line.split("^")
                        person = {"firstname":_person[2]}
                        person["lastname"] = _person[5]
                        person["jobtitle"] = _person[4]
                        person["organization"] = _person[6]
                        person["id"] = _person[3]
                        person["email"] = _person[0]
                        person["phone"] = _person[8]

                        persons.append(person)
                else:
                        linecount=linecount+1
        fh.close()
        return persons

def _not(value,nottest):
        if nottest==True and value == 1:
                return 0
        elif nottest==True and value == 0:
                return 1
        else:
                return value

def _testmatch(jobtitle,constraint,_match,testdepth,nottest=False):
        if jobtitle.find(constraint)==-1:
                _match=_match*_not(0,nottest)
        else:
                _match=_match*_not(1,nottest)
        #if DEBUG!=False:
        #        print(" "*testdepth)
        #        print("not_test_type="+str(nottest) + " jobtitle=" + jobtitle + " constraint=" + constraint + " matchstr=" + str(_match))
        return _match

def _update_summary(summary,value,_id,category):
        
        if value in summary[category].keys():
                _tmp = summary[category][value]
                _tmp.append(_id)
                summary[category][value] = _tmp
        else:
                summary[category][value] = [_id]
        return summary

t1 = datetime.datetime.now()
numtests=0


if "debug" in locals().keys():
        DEBUG=locals()["debug"]
else:
        DEBUG=False
        
if "outputfile" in locals().keys():
        outputfile=locals()["outputfile"]
else:
        outputfile="output.txt"


persons = get_persons(locals()["personsfile"])
clients = get_clients(locals()["clientsfile"])
rulesets =get_rulesets(locals()["rulesfile"],delimiter)
count =0



results={}
summary={"Job Type":{},"Sub Department":{},"Seniority":{},"Department":{}}
for _person in persons:
        _jobtitle= _person["jobtitle"]
        _id=_person["id"]

        if _id=="":
                id=time.time() * 1000

        person_results={"debug":[]}
        for _ruleset in rulesets.keys():
                for _rule in rulesets[_ruleset]:
                        constraints = _rule[3].split("$$")
                        category=_rule[0]
                        value=_rule[1]
                        match=1

                        
                        # for debugging
                        #print(_person["jobtitle"],_person["id"],constraints, category, value)
                        
                        
                        for i in range(0,len(constraints)):
                                _constraint = constraints[i]
                                if _constraint.find("!")==-1:
                                        match = _testmatch(_jobtitle,_constraint,match,i)
                                else:
                                        _constraint=_constraint.replace("!","")
                                        match = _testmatch(_jobtitle,_constraint,match,i,nottest=True)

                        numtests=numtests+1
                        if match==1:
                                if DEBUG!=False:
                                        #print(_ruleset + ":" + _jobtitle + " MATCH [" + ",".join(_rule)+"]")
                                        person_results["debug"].append(_ruleset + ":" + _jobtitle + " MATCH [" + ",".join(_rule)+"]")
                                person_results[category]=value
                                _update_summary(summary,value,_id,category)
                                break

                if match==0:
                        if DEBUG!=False:
                                #print(_ruleset + ":" + _jobtitle + " NOMATCH [" + ",".join(_rule)+"]")
                                person_results["debug"].append(_ruleset + ":" + _jobtitle + " NOMATCH [" + ",".join(_rule)+"]")
                        person_results[category]="UNKNOWN"
                        _update_summary(summary,value,_id,category)
                else:
                        pass

        count = count +1
        results[_id]={"results": person_results, 
              "organization":_person["organization"], 
              "jobtitle" : _jobtitle, 
              "email" : _person["email"], 
              "phone" : _person["phone"], 
              "firstname":_person["firstname"], 
              "lastname":_person["lastname"]}

sys.stderr.write(str(count))
print_results(results,clients)
t2 = datetime.datetime.now()
delta = t2-t1
sys.stderr.write("number of tests performed: " + str(numtests) + " in " +str(delta) + "secs\n")

import pickle
with open('summary.pickle', 'wb') as f:
        pickle.dump(summary, f, pickle.HIGHEST_PROTOCOL)
with open('results.pickle', 'wb') as f:
        pickle.dump(results, f, pickle.HIGHEST_PROTOCOL)