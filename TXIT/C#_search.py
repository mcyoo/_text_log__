import os
import codecs
import sys

#파일안에 hostname을 추출하는 함수
def name(file,filename):
	name = 0
	file.seek(0)
	while True:
		line=file.readline()
		hop = line.count('#')
		if not line:
			line = filename
			break
		if hop == 0 or hop > 5: continue
		else :
			temp = line
			for i in range(0,hop):
				name+=temp.find('#')
				name+=1
				temp=line[name:]
				i+=1

			line=line[:name-1]
			break

	return line

#function_input 값을 file 안에서 찾아서 match 가 되면 return 하는 함수
def search_input_match_return(file,function_input,filename):

	hostname=name(file,filename)
	file.seek(0)
	while True:
		line = file.readline()
		tem1=line.find(function_input)
		
		if tem1 >= 0:
			return hostname
		if not line: 
			break

#function_input 값을 file 안에서 찾아서 match 가 안되는 것을 return 하는 함수
def search_input_notmatch_return(file,function_input,filename):

	hostname=name(file,filename)
	file.seek(0)
	while True:
		line = file.readline()
		tem1=line.find(function_input)
		
		if tem1 >= 0:
			break
		if not line:
			return hostname

#toggle 값이 true 이면, user_input 이 match 되는걸 찾고, 아니면 match가 안될걸 찾는다.
def main_search(user_input,pathnote,toggle):
	search_list = []
	search_list.clear()
	
	for (path, dir, files) in os.walk(pathnote):
	   
		for filename in files:
			ext = os.path.splitext(filename)[-1]
			if ext == '.txt' or ext == '.log' :
				print("%s/%s" % (path, filename))
	
				f = codecs.open("%s/%s" % (path, filename), 'r', "utf-8-sig", errors='ignore')#utf-8-sig
				
				if toggle == '0':
					search_list.append(search_input_match_return(f,user_input,filename))
				else:
					search_list.append(search_input_notmatch_return(f,user_input,filename))

	f.close()
	
	#리스트에 None 값 제거
	while None in search_list:
		search_list.remove(None)
	
	return search_list

if __name__ == "__main__":

	args = sys.argv[1:]
	path = args[0]
	search = args[1]
	toggle = args[2]

	search_list = main_search(search,path,toggle)
	
	for i in search_list:
		print(i)

	print("총 " + str(len(search_list)) + "개 찾았습니다.")


