# Creation of Road
def createRoad(all_street,street,street_detail,city_map_dict):
    all_street.append(street)
    newStreet = []

    for i in range(0,len(street_detail[1:])):
        x = street_detail[i]
        y = street_detail[i+1]
        if x[-1] == ')' and  y[-1] == ')':
            newStreet.append( x + '-->'+ y )
        else:
            sys.stderr.write("Enter All Coordinates within parenthesis")    
    city_map_dict[street] = newStreet
    return city_map_dict

# Change Road
def changeRoad(all_street,street,street_detail,city_map_dict):
    newStreet = []

    for i in range(0,len(street_detail[1:])):
        x = street_detail[i]
        y = street_detail[i+1]
        if x[-1] == ')' and  y[-1] == ')':
            newStreet.append( x + '-->'+ y )
        else:
            sys.stderr.write("Enter All Coordinates within parenthesis")    
    city_map_dict[street] = newStreet    
    return city_map_dict
                
# Deletion of Road
def deleteRoad(city_map_dict,street,city_map,all_street):
    rindex = all_street.index(street)
    
    # del city_map[rindex]
    del city_map_dict[street]
    all_street.remove(street)
    pass

#Updating Dictionary
def updateDict(city_map,city_map_dict,street,all_street):
    uindex = all_street.index(street)
    # for lis in city_map[:]:
    city_map_dict[street] = city_map[uindex]

#To Calculate keep the Vertices key in same order after deletion using a copy of prev
def verticesOrder(vertices,vkeys,vdict,vertices1):
    # ikey = 0
    # if bool(vertices) == False: 
    #     ikey = 0
    #     for key in vkeys:
    #         vertices[ikey] =ast.literal_eval(key)
    #         ikey =ikey + 1
    #     for value in vdict.values():
    #         for item in range(0,len(value)):
    #             if not ast.literal_eval(value[item]) in vertices.values():
    #                 vertices[ikey] =  ast.literal_eval(value[item])
    #                 ikey = ikey + 1
    if not vdict == {}:
        vertices.clear()
        ikey = 0
        for key in vkeys:
            vertices[ikey] =ast.literal_eval(key)
            ikey =ikey + 1
        for value in vdict.values():
            for item in range(0,len(value)):
                if not ast.literal_eval(value[item]) in vertices.values():
                    vertices[ikey] =  ast.literal_eval(value[item])
                    ikey = ikey + 1    
        # ikey = 0
        # for key in vkeys:
        #     vertices1[ikey] = ast.literal_eval(key)
        #     ikey =ikey + 1
        # for value in vdict.values():
        #     for item in range(0,len(value)):
        #         if ast.literal_eval(value[item]) not in vertices1.values():
        #             vertices1[ikey] = ast.literal_eval(value[item])
        #             ikey = ikey + 1
        # vertices = vertices1.copy()    
        # return vertices    

        
        # for keyold,oldvalue in vertices.items():
        #     flag = 0
        #     for newvalue in vertices1.items():
        #         if oldvalue == newvalue[1]:
        #             break
        #         else:
        #             flag = flag + 1
        #             if flag == len(vertices1):
        #                 del vertices[keyold]
        
        # if vertices == {}:

        #     for key,newvalue1 in vertices1.items():
        #         vertices[key]= newvalue1

        # else:
        #     for newvalue1 in vertices1.items():
        #         flag = 0
        #         for oldvalue1 in vertices.items():
        #             if oldvalue1[1] == newvalue1[1]:
        #                 break
        #             else:
        #                 flag = flag + 1
        #                 if flag == len(vertices):
        #                     keyv = oldvalue1[0] + 1
        #                     vertices[keyv] = newvalue1[1]
        
    else:
        vertices.clear()
        

def makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V):
    l1 = base
    l2 = compar    
    temp = []                            
    if l1 not in lineeqn:
        lineeqn.append(l1)
    if l2 not in lineeqn:
        lineeqn.append(l2)

    if V not in intercordinal_keys:
        intercordinal_keys.append(V)
        temp.append(l1)
        temp.append(l2)
        intercordinal[V]= temp[:]
    else:
        if not l1 in intercordinal[V]:
            intercordinal[V].append(l1)    
        if not l2 in intercordinal[V]:
            intercordinal[V].append(l2)    

#Calculate Intersection points
def find_intersection(x1,y1,x2,y2,x3,y3,x4,y4) :

    s10_x = x2 - x1
    s10_y = y2 - y1
    s32_x = x4 - x3
    s32_y = y4 - y3
    denom = s10_x * s32_y - s32_x * s10_y

    if denom == 0 : 
        x,y = '',''
        return (x,y) # collinear
    denom_is_positive = denom > 0
    s02_x = x1 - x3
    s02_y = y1 - y3
    s_numer = s10_x * s02_y - s10_y * s02_x
    if (s_numer < 0) == denom_is_positive : 
        x,y = '',''
        return (x,y) # no collision
    t_numer = s32_x * s02_y - s32_y * s02_x
    if (t_numer < 0) == denom_is_positive :
        x,y = '',''
        return (x,y) # no collision
    if (s_numer > denom) == denom_is_positive or (t_numer > denom) == denom_is_positive : 
        x,y = '',''
        return (x,y) # no collision
    t = t_numer / denom # collision detected
    x,y = x1 + (t * s10_x), y1 + (t * s10_y) 
    return x,y
            
# Calculating Vertices
def calculateVertices(city_map_dict,all_street,vertices,vdict):

    
    #Local Declaraions
    vlist = []
    vkeys = []
    vertices1 = {}
    intercordinal = {}
    lineeqn = []
    intercordinal_keys = []

    for i in range(0,len(all_street)-1):           #loop through all the streets
        key1 = all_street[i]                       #Take one street and all its directions
        for base in city_map_dict[key1][:]:
            j = i + 1
            line1 = re.findall(r'(-?\d+\.?\d*)',base)                #Base Segment to Compare
        
            for k in range (j,len(all_street)):                   #Take the next street in the list and get all its directions
                for compar in city_map_dict[all_street[k]][:]: #Take the next street's each coord to compare with initial street
                    line2 = re.findall(r'(-?\d+\.?\d*)',compar)       #Segment to Compare with
                    
                    ##Coordinates x,y values
                    x1, y1 = float(line1[0]), float(line1[1])
                    x2, y2 = float(line1[2]), float(line1[3])
                    x3, y3 = float(line2[0]), float(line2[1])
                    x4, y4 = float(line2[2]), float(line2[3])

    #Case1: 5 cases Check if they have common points meaning they end and another street starts
                    #or if the end and then new street beings (overlapping)
                    if x1 == x3 and y1 == y3:            
                        V = '(' + str(x1) + ',' + str(y1) +')'     #If yes take the common point as intersection
                        if V not in vkeys:            #Check if its a new i.point

                        #Vertex Calculation Dict Purpose    
                            vkeys.append(V)            #Append it to vkey(storing all int points)
                            vlist.append('(' + str(x2) + ',' + str(y2) +')')
                            vlist.append('(' + str(x4) + ',' + str(y4) +')')                
                            vdict[V] = vlist[:]        #Append to Vdict by storing i.point as key and others as points
                            del vlist[:]            #causing the i.point    
                        
                        #Edge- Intersection Lines Dict purpose
                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)

                        else: #If it exists already append only the new coordinate causing I.point to vidct's value in the same key

                        #Vertex Duplication Reduction purpose    
                            tempcoord = '(' + str(x2) + ',' + str(y2) +')'
                            if tempcoord not in vdict[V]:
                                vdict[V].append(tempcoord)        
                            tempcoord = '(' + str(x4) + ',' + str(y4) +')'
                            if tempcoord not in vdict[V]:
                                vdict[V].append(tempcoord)
                        
                        #Edge Calculation Dict purpose
                            #Edge- Intersection Lines Dict purpose
                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)

    #Case2:
                    elif x1 == x4 and y1 == y4:
                        V = '(' + str(x1) + ',' + str(y1) +')' 
                        if V not in vkeys:
                    
                            vkeys.append(V)
                            vlist.append('(' + str(x2) + ',' + str(y2) +')')
                            vlist.append('(' + str(x3) + ',' + str(y3) +')')
                            vdict[V] = vlist[:]
                            del vlist[:]    
                        #Edge- Intersection Lines Dict purpose
                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)



                        else:
                            tempcoord = '(' + str(x2) + ',' + str(y2) +')'
                            if tempcoord not in vdict[V]:
                                vdict[V].append(tempcoord)        
                            tempcoord = '(' + str(x3) + ',' + str(y3) +')'
                            if tempcoord not in vdict[V]:
                                vdict[V].append(tempcoord)
                        #Edge- Intersection Lines Dict purpose
                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)


    #Case3:                
                    elif x2 == x3 and y2 == y3:
                        V = '(' + str(x2) + ',' + str(y2) +')' 
                        if V not in vkeys:
                    
                            vkeys.append(V)
                            vlist.append('(' + str(x1) + ',' + str(y1) +')')
                            vlist.append('(' + str(x4) + ',' + str(y4) +')')
                            vdict[V] = vlist[:]
                            del vlist[:]

                            #Edge- Intersection Lines Dict purpose
                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)

                        else:
                            tempcoord = '(' + str(x1) + ',' + str(y1) +')'
                            if tempcoord not in vdict[V]:
                                vdict[V].append(tempcoord)        
                            tempcoord = '(' + str(x4) + ',' + str(y4) +')'
                            if tempcoord not in vdict[V]:
                                vdict[V].append(tempcoord)
                            #Edge- Intersection Lines Dict purpose
                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)

    #Case4                
                    elif x2 == x4 and y2 == y4:
                        V = '(' + str(x2) + ',' + str(y2) +')' 
                        if V not in vkeys:
                    
                            vkeys.append(V)
                            vlist.append('(' + str(x1) + ',' + str(y1) +')')
                            vlist.append('(' + str(x3) + ',' + str(y3) +')')
                            vdict[V] = vlist[:]
                            del vlist[:]
                        #Edge- Intersection Lines Dict purpose
                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)

                        else:
                            tempcoord = '(' + str(x1) + ',' + str(y1) +')'
                            if tempcoord not in vdict[V]:
                                vdict[V].append(tempcoord)        
                            tempcoord = '(' + str(x3) + ',' + str(y3) +')'
                            if tempcoord not in vdict[V]:
                                vdict[V].append(tempcoord)
                            #Edge- Intersection Lines Dict purpose
                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)

                        
    #Case5:Find I.point using slope    function            
                    else:

                        xcoor,ycoor = find_intersection(x1,y1,x2,y2,x3,y3,x4,y4)
                        if not xcoor == '' and not ycoor == '':
                            V = '(' + str(xcoor) + ',' + str(ycoor) +')'

                            #If intersection found then and its unique
                            if V not in vkeys:
                        
                                #Vertice Calculation Dict Purpose
                                vkeys.append(V)
                                vlist.append('(' + str(x1) + ',' + str(y1) +')')
                                vlist.append('(' + str(x2) + ',' + str(y2) +')')
                                vlist.append('(' + str(x3) + ',' + str(y3) +')')
                                vlist.append('(' + str(x4) + ',' + str(y4) +')')
                                vdict[V] = vlist[:]
                                del vlist[:]

                                #Edge- Intersection Lines Dict purpose
                                makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)


                            else:#If intersection found and its already present
                                
                                #Vertices Calculation purpose (Removing duplicates)
                                tempcoord = '(' + str(x1) + ',' + str(y1) +')'
                                if tempcoord not in vdict[V]:
                                    vdict[V].append(tempcoord)
                                tempcoord = '(' + str(x2) + ',' + str(y2) +')'
                                if tempcoord not in vdict[V]:
                                    vdict[V].append(tempcoord)
                                tempcoord = '(' + str(x3) + ',' + str(y3) +')'
                                if tempcoord not in vdict[V]:
                                    vdict[V].append(tempcoord)
                                tempcoord = '(' + str(x4) + ',' + str(y4) +')'
                                if tempcoord not in vdict[V]:
                                    vdict[V].append(tempcoord)    

                                #Edge- Intersection Lines Dict purpose
                                makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)
                        else:
                                vpointdist = []
                                i = 0
                                vpointdist = point_exists_InLine(x1,y1,x2,y2,x3,y3,x4,y4,vpointdist)
                                if not vpointdist == []:
                                    for eindex in range(0,len(vpointdist)-1,2):    
                                        V = '(' + str(vpointdist[eindex]) + ',' + str(vpointdist[eindex+1]) +')'

                                        #If intersection found then and its unique
                                        if V not in vkeys:
                                    
                                            #Vertice Calculation Dict Purpose
                                            vkeys.append(V)
                                            vlist.append('(' + str(x1) + ',' + str(y1) +')')
                                            vlist.append('(' + str(x2) + ',' + str(y2) +')')
                                            vlist.append('(' + str(x3) + ',' + str(y3) +')')
                                            vlist.append('(' + str(x4) + ',' + str(y4) +')')
                                            vdict[V] = vlist[:]
                                            del vlist[:]

                                            #Edge- Intersection Lines Dict purpose
                                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)


                                        else:#If intersection found and its already present
                                            
                                            #Vertices Calculation purpose (Removing duplicates)
                                            tempcoord = '(' + str(x1) + ',' + str(y1) +')'
                                            if tempcoord not in vdict[V]:
                                                vdict[V].append(tempcoord)
                                            tempcoord = '(' + str(x2) + ',' + str(y2) +')'
                                            if tempcoord not in vdict[V]:
                                                vdict[V].append(tempcoord)
                                            tempcoord = '(' + str(x3) + ',' + str(y3) +')'
                                            if tempcoord not in vdict[V]:
                                                vdict[V].append(tempcoord)
                                            tempcoord = '(' + str(x4) + ',' + str(y4) +')'
                                            if tempcoord not in vdict[V]:
                                                vdict[V].append(tempcoord)    

                                            #Edge- Intersection Lines Dict purpose
                                            makeIntlineDict(base,compar,intercordinal,intercordinal_keys,lineeqn,V)




    verticesOrder(vertices,vkeys,vdict,vertices1) #Vertice Calculation and order
    return (intercordinal,lineeqn,intercordinal_keys)
    
#CAlculate EDges        
def calculatEdges(intercordinal,vdict,vertices,lineeqn):
    intersecs = []
    iodistance = []
    edges = []
    for line in lineeqn:
        for k,v in intercordinal.items():
            if line in v:
                if k not in intersecs:
                    intersecs.append(k)
        a1,a2,a3,a4 = re.findall(r'(-?[+-]?\d*\.?\d+)',line)
        if len(intersecs) == 1:
            ck1 = '('+ a1 + ',' + a2 + ')'
            ck2 = '('+ a3 + ',' + a4 + ')'
            ck1 = ast.literal_eval(ck1)
            ck2 = ast.literal_eval(ck2)
            if not str(vertices.keys()[vertices.values().index(ast.literal_eval(intersecs[0]))]) == str( vertices.keys()[vertices.values().index(ck1)]):
                e1 = '<' + str(vertices.keys()[vertices.values().index(ast.literal_eval(intersecs[0]))]) + ',' + str( vertices.keys()[vertices.values().index(ck1)]) + '>'
                if e1 not in edges:edges.append(e1)
            if not str(vertices.keys()[vertices.values().index(ast.literal_eval(intersecs[0]))]) == str( vertices.keys()[vertices.values().index(ck2)]):
                e1 = '<' + str(vertices.keys()[vertices.values().index(ast.literal_eval(intersecs[0]))]) + ',' + str( vertices.keys()[vertices.values().index(ck2)]) + '>'
                if e1 not in edges: edges.append(e1)
            del intersecs[:]
            ck1,ck2 = '',''
        
        if len(intersecs) > 2:

            pointointdistance=[]
            endpointdistance = []
            for echi in range(0,len(intersecs)):
                c1,c2 = re.findall(r'(-?[+-]?\d*\.?\d+)',intersecs[echi])
                pidist = math.sqrt( (float(c2) - float(a2))**2 + (float(c1) - float(a1))**2 )
                pointointdistance.append(pidist)
                endpidist = math.sqrt( (float(c2) - float(a4))**2 + (float(c1) - float(a3))**2 )
                endpointdistance.append(endpidist)
                
            f1,f2 = re.findall(r'(-?[+-]?\d*\.?\d+)',(intersecs[pointointdistance.index(min(pointointdistance))]))
            g1,g2 = re.findall(r'(-?[+-]?\d*\.?\d+)',(intersecs[endpointdistance.index(min(endpointdistance))]))
            newintersecs = intersecs[:]
            
            ck1 = '('+ f1 + ',' + f2 + ')'    
            ck2 = '('+ a1 + ',' + a2 + ')'
            reference = ck1
            newintersecs.remove(ck1)
            ck1 = ast.literal_eval(ck1)
            ck2 = ast.literal_eval(ck2)

            if not str(vertices.keys()[vertices.values().index(ck1)]) == str( vertices.keys()[vertices.values().index(ck2)]):
                e1 = '<' + str(vertices.keys()[vertices.values().index(ck1)]) +  ',' + str( vertices.keys()[vertices.values().index(ck2)]) + '>'
                if e1 not in edges: edges.append(e1)

            ck1 = ''
            ck2 = ''

            ck1 = '('+ g1 + ',' + g2 + ')'    
            ck2 = '('+ a3 + ',' + a4 + ')'
            # newintersecs.remove(ck1)
            # endrefer = ck1
            ck1 = ast.literal_eval(ck1)
            ck2 = ast.literal_eval(ck2)

            if not str(vertices.keys()[vertices.values().index(ck1)]) == str( vertices.keys()[vertices.values().index(ck2)]):
                e1 = '<' + str(vertices.keys()[vertices.values().index(ck1)]) + ','+ str( vertices.keys()[vertices.values().index(ck2)]) + '>'
                if e1 not in edges: edges.append(e1)

            ck1 = ''
            ck2 = ''

            for inte in newintersecs[:]:
                # if not len(newintersecs) == 1:
                for ke in range(0,len(newintersecs)):
                    b1,b2 = re.findall(r'(-?[+-]?\d*\.?\d+)',reference)            
                    b3,b4 = re.findall(r'(-?[+-]?\d*\.?\d+)',newintersecs[ke])
                    dist = math.sqrt ( (float(b4) - float(b2))**2 + (float(b3) - float(b1))**2 )
                    iodistance.append(dist)
                if not iodistance == []    :
                    kindex = iodistance.index(min(iodistance))
                    b = re.findall(r'(-?[+-]?\d*\.?\d+)',(newintersecs[kindex]))
                    b3 = b[0]
                    b4 = b[1]
                    ck1 = reference    
                    ck2 = '('+ b3 + ',' + b4 + ')'
                    ck1 = ast.literal_eval(ck1)
                    ck2 = ast.literal_eval(ck2)
                    if not str(vertices.keys()[vertices.values().index(ck1)]) == str( vertices.keys()[vertices.values().index(ck2)]):
                        e1 = '<' + str(vertices.keys()[vertices.values().index(ck1)]) + ',' + str( vertices.keys()[vertices.values().index(ck2)]) + '>'
                        if e1 not in edges:edges.append(e1)
                    reference = '('+ b3 + ',' + b4 + ')'
                    newintersecs.pop(kindex)
                    ck1,ck2 = '',''
                    iodistance = []
                    # else:
                    #     if not str(vertices.keys()[vertices.values().index(endrefer)]) == str(vertices.keys()[vertices.values().index(newintersecs[0])]):
                    #         e1 = '<' + str(vertices.keys()[vertices.values().index(endrefer)]) + ',' + str(vertices.keys()[vertices.values().index(newintersecs[0])]) + '>'
                    #     if e1 not in edges:edges.append(e1)
                else:
                    break


        elif len(intersecs) == 2:
                    twodistance = []
                    c1,c2 = re.findall(r'(-?[+-]?\d*\.?\d+)',intersecs[0])
                    twodist = math.sqrt( (float(c2) - float(a2))**2 + (float(c1) - float(a1))**2 )
                    twodistance.append(twodist)
                    twodist = math.sqrt( (float(c2) - float(a4))**2 + (float(c1) - float(a3))**2 )
                    twodistance.append(twodist)

                    if twodistance.index(min(twodistance)) == 0:
                        fr1  = ast.literal_eval(intersecs[0])
                        end1 = ast.literal_eval(intersecs[1])
                    else:
                        fr1  = ast.literal_eval(intersecs[1])
                        end1 = ast.literal_eval(intersecs[0])
                    ck1 = '('+ a1 + ',' + a2 + ')'
                    ck2 = '('+ a3 + ',' + a4 + ')'
                    ck1 = ast.literal_eval(ck1)
                    ck2 = ast.literal_eval(ck2)
                    
                    if not str(vertices.keys()[vertices.values().index(fr1)]) == str( vertices.keys()[vertices.values().index(ck1)]):
                        e1 = '<' + str(vertices.keys()[vertices.values().index(fr1)]) + ',' + str( vertices.keys()[vertices.values().index(ck1)]) + '>'
                        if e1 not in edges: edges.append(e1)
                    if not str(vertices.keys()[vertices.values().index(end1)]) == str( vertices.keys()[vertices.values().index(ck2)]):
                        e1 = '<' + str(vertices.keys()[vertices.values().index(end1)]) + ',' + str( vertices.keys()[vertices.values().index(ck2)]) + '>'
                        if e1 not in edges: edges.append(e1)

                    i1 = ast.literal_eval(intersecs[0])
                    i2 = ast.literal_eval(intersecs[1])
                    if not str(vertices.keys()[vertices.values().index(i1)]) == str( vertices.keys()[vertices.values().index(i2)]):
                        e1 = '<' + str(vertices.keys()[vertices.values().index(i1)]) + ',' + str( vertices.keys()[vertices.values().index(i2)]) + '>'
                        if e1 not in edges: edges.append(e1)
        intersecs = []

    # U=0
    # for i in range(0,len(edges)):
    #     edges[i] = ast.literal_eval(edges[i])
    return(edges)

def point_exists_InLine(x1,y1,x2,y2,x3,y3,x4,y4,vpointdist):

    if not (x4 - x3) == 0 and not (x2 - x1) == 0:
        slope1 = (y2 - y1) / (x2 - x1)
        slope2 = (y4-y3) / (x4-x3)
        if slope1 == slope2:

            Intercept1 = y1 - (slope1 * x1)
            Intercept2 = y3 - (slope2 * x3)

                ##Check if X1,Y1 lies in Equation2
            Interceptcheck = y1 - (slope2 * x1)
            if Interceptcheck == Intercept2:
            
                # if (x2-x1) == 0 or  (x4-x3) == 0:
                if min(x1,x2) <= min(x3,x4) and max(x1,x2) >= max(x3,x4) and min(y1,y2) <= min(y3,y4) and max(y1,y2) >= max(y3,y4):

                        vpointdist.append(x3)
                        vpointdist.append(y3)
                        vpointdist.append(x4)
                        vpointdist.append(y4)
                elif min(x3,x4) <= min(x1,y1) and max(x3,x4) >= max(x1,x2) and min(y3,y4) <= min(y1,y2) and max(y3,y4) >= max(y1,y2):

                        vpointdist.append(x1)
                        vpointdist.append(y1)
                        vpointdist.append(x2)
                        vpointdist.append(y2)
    else:

        if x1 == x2 == x3 == x4:
            if min(x1,x2) <= min(x3,x4) and max(x1,x2) >= max(x3,x4) and min(y1,y2) <= min(y3,y4) and max(y1,y2) >= max(y3,y4):

                    vpointdist.append(x3)
                    vpointdist.append(y3)
                    vpointdist.append(x4)
                    vpointdist.append(y4)
            elif min(x3,x4) <= min(x1,y1) and max(x3,x4) >= max(x1,x2) and min(y3,y4) <= min(y1,y2) and max(y3,y4) >= max(y1,y2):

                    vpointdist.append(x1)
                    vpointdist.append(y1)
                    vpointdist.append(x2)
                    vpointdist.append(y2)

        # elif x2-x1 == 0 and ( x1 == x3 or x1 == x4):
        #     vpointdist.append(x1)
        #     slope2 = (y4-y3) / (x4-x3)
        #     Inter = y3 - (slope2 * x3)
        #     yn = (slope2*x1) + Inter
        #     vpointdist.append(yn)

        # elif x3-x4 == 0 and (x4 == x1 or x4 == x2):
        #     vpointdist.append(x3)
        #     slope2 = (y2-y1) / (x2-x1)
        #     Inter = y2 - (slope2 * x2)
        #     yn = (slope2*x3) + Inter
        #     vpointdist.append(yn)

    return(vpointdist)
