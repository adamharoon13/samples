import sys
import re
from calcu import*



class ParseException(Exception):
    def __init__(self, value):
        self._value = value

    def __str__(self):
        return str(self._value)

def main():
    run(sys.stdin, sys.stdout, sys.stderr)

def run(input, out, err):
    while True:
        
        # out.write("\nEnter the Function, Street Name and the Coordinates separated with Space and press enter or type exit to leave:")
        s = input.readline()
        s = s.replace('\n','')

        if s == '':
                break

        elif s == "DONE":
                out.write('DONE')
                out.flush()  
                break
        try:
                parse_line(s,input,out,err)
        except ParseException as ex:
                    err.write("Error: {0}\n".format(ex))


def parse_line(s,input,out,err):
    if not s.startswith('a') and not s.startswith('c') and not s.startswith('g') and not s.startswith('r'):
        raise ParseException("does not start with 'add' or 'mod' or 'rm' or 'gg'")
    else:
        s = s.replace("\n",'')
        data = s.split(" ",1)
        if len(data[0]) == 1:
            func = data[0] # To Determine Function
            
            if s.startswith('g') == False:
                    if len(data) > 1:
                        splitup = data[1].split('"(',1)
                        if len(splitup) > 1 and not s.startswith('r') == 'True':
                            raise ParseException("Missing Space between street and coordinates")
                        street_temp = data[1].split('"',2)
                        if len(street_temp) == 3:
                            street_name = street_temp[1].lower()
                            if not re.match(r'.[a-zA-Z\s]*$',street_name):
                                raise ParseException("No special character allowed in street name")
                            street = street_name.replace(' ','_').replace('"',"")
                            
                            street_temp[2] = street_temp[2].replace(" ",'')
                            street_temp[2] = street_temp[2].replace("\n",'')
                            rawcoord = len(street_temp[2])
                            check_pair_coord =re.findall('\s*[-]*[0-9]+\.?[0-9]*\s*',street_temp[2])

                            if not len(check_pair_coord) % 2 == 0:
                                raise ParseException("Missing Pair X,Y of Coordinate values")

                            parsed_coord = re.findall(r'\(.*?\)',street_temp[2])
                            op_count = street_temp[2].count('(')
                            cp_count = street_temp[2].count(')')

                            temp = 0
                            for i in range(0,len(parsed_coord)):
                                temp = temp + len(parsed_coord[i])
                                
                            if not rawcoord == temp or not op_count == cp_count:
                                raise ParseException("Missing Open/Closure Parenthesis")
                            
                            street_detail = parsed_coord
                        elif len(street_temp) == 2:
                            if not street_temp[0] == '"' and not street_temp[-1] == '"':
                                raise ParseException("Please provide Street in Quotes")
                            street_name = street_temp[1].lower()
                            if not re.match(r'.[a-zA-Z\s]*$',street_name):
                                raise ParseException("No special character allowed in street name")
                            street = street_name.replace(' ','_').replace('"',"")
                            street = street.replace("\n",'')
                        else: raise ParseException("Insufficient Data")    
                        
                        if street not in all_street:
                            if func == 'add':    

                                createRoad(all_street,street,street_detail,city_map_dict)
                                

                            if func == 'mod':
                                        raise ParseException("'mod' specified for a street that does not exist")
                                        pass
                            elif func == 'rm':
                                        raise ParseException("'rm' specified for a street that does not exist")
                                        pass
                        else:
                            if func == 'add':    
                                raise ParseException("'add' specified for a street that already exists use 'mod'")
                                pass
                            elif func == 'mod':
                                changeRoad(all_street,street,street_detail,city_map_dict)
                            elif func == 'rm':
                                deleteRoad(city_map_dict,street,city_map,all_street)        
                    else:
                        raise ParseException("Insufficient Data")            
            else:
                    intercordinal,lineeqn,intercordinal_keys =    calculateVertices(city_map_dict,all_street,vertices,vdict) #Calculate Vertices
                    edges = calculatEdges(intercordinal,vdict,vertices,lineeqn) #Calculate Edges
                    vstring = vertices
            #          out.write("V = {\n")
            #         for k, v in vertices.items():
            #             vround = (round(v[0],2),round(v[1],2))                
            #             vround= str(vround)
            #             vround= vround.replace(" ",'')        
            #             out.write("  {}:  {}\n".format(k, vround))
                        
            #         out.write("}")
                    lenv = len(vertices)
                    out.write(format('V') +' '+ format(lenv))
                    out.flush()
                    out.write("\nE {" + ",".join(edges) + "}\n")
                    out.flush()
                    vdict.clear()
                    lineeqn = []
                    edges = []
                    intercordinal_keys = []
                    intercordinal.clear()
        else:
            raise ParseException("Please Seperate Function,Street and Coordinates with white spaces")    

if __name__ == '__main__':
    #Declare Global Variables
    all_street = [] # List to store all road names
    city_map   = [] # List to store Names,Coordinates
    city_map_dict   = {}
    lines= []
    vertices = {}
    vdict = {}
    edges = []
    intercordinal = {}
    intercordinal_keys = []
    lineeqn = []
    V = ''
    main()