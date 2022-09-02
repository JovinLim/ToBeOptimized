from operator import indexOf
import numpy as np
import pandas as pd
import math as m
import random as ran
from typing import Counter
import matplotlib.pyplot as plt
import copy
import numpy as np
import sys
import queue
from pymoo.core.problem import ElementwiseProblem
from pymoo.core.problem import Problem
from pymoo.problems.functional import FunctionalProblem
from pymoo.optimize import minimize
from pymoo.factory import get_termination
from pymoo.algorithms.moo.nsga2 import NSGA2
from pymoo.factory import get_algorithm, get_crossover, get_mutation, get_sampling
from pymoo.decomposition.asf import ASF
from pymoo.core.evaluator import Evaluator
from pymoo.core.population import Population
from pymoo.operators.crossover.sbx import SBX
from pymoo.operators.mutation.pm import PM
from pymoo.operators.repair.rounding import RoundingRepair
import randfacts

print ("Hello from the TBO Team! Please wait for our optimisation algorithm to complete...")
print (" \n \n " )
filepath = 'C:\\temp\\SOA_Copy.xlsx'
filepath2 = 'C:\\temp\\param.xlsx'
writepath = 'C:\\temp\\Output.xlsx'

matrix_setting = {'XY_Rotate': None, 'XY_Mirror': None, 'Entrance_Dir': None, 'New_Ent_Dir': None, 'Ent_Pt': None}

df = pd.read_excel(filepath, sheet_name = 'Orthopaedic', header = 1,index_col=0, usecols=['Room Type','Unit Area/sqm','Quantity'])
test_dict = df.to_dict()
SOA_dict = test_dict['Quantity']
print(SOA_dict)
print (" \n \n \n" )

df2 = pd.read_excel(filepath, sheet_name = 'Orthopaedic', header = 1, index_col = 0, usecols = ['Axis','1m Matrix Grid'])
test_dict2 = df2.to_dict()
xy_grid = test_dict2['1m Matrix Grid']


#Determine Orientation from entrance points in excel file
df3 = pd.read_excel(filepath, sheet_name = 'Orthopaedic', header = 2, index_col = None, usecols = 'J',nrows=0)
excel_cor = df3.columns.values[0]

df4 = pd.read_excel(filepath2, sheet_name = 'Objectives', index_col = 0, usecols = ['Obj_Name','Y/N'])
test_dict4 = df4.to_dict()
TBO_Obj = test_dict4['Y/N']


df5 = pd.read_excel(filepath2, sheet_name = 'Parameters', index_col = 0, usecols = ['Param_Name','Y/N'])
test_dict5 = df5.to_dict()
TBO_Param = test_dict5['Y/N']


#parse string from excel to get index of corridor abutting SOC
def readstring(ex_str,y_ax,x_ax):
  y_output = []
  x_output = []
  temp = ""
  for i in ex_str:
    count = 0
    if i.isdigit() == True:
      temp = temp + i
    elif i == ',':
      y_output.append(int(temp))
      temp = ""
    elif i == '|':
      x_output.append(int(temp))
      temp = ""
    else:
      print('Error: Cannot parse string')
  Ent_pts = []
  if len(y_output) == len(x_output):
    for j in range(len(y_output)):
      Ent_pts.append([y_output[j],x_output[j]])
  else:
    print('Error: y_output does not match x_output')
  #Group corridor ID by y and x value in order to determine orientation of access
  y_group = []
  x_group = []
  for i in range(y_ax):
    num_county = 0
    num_countx = 0
    for j in range(len(Ent_pts)):
      if Ent_pts[j][0] == i:
        num_county += 1
      if Ent_pts[j][1] == i:
        num_countx += 1
    y_group.append(num_county)
    x_group.append(num_countx)
  #In the order of East, North, West, South 
  dir_count = [x_group[-1],y_group[-1],x_group[0],y_group[0]] 
  return dir_count,Ent_pts
    
dir_count,Ent_pts = readstring(excel_cor,int(xy_grid['y_ax']),int(xy_grid['x_ax']))

matrix_setting['Entrance_Dir'] = dir_count.index(max(dir_count))

#Adjust matrix so that SOC generated starts from the South, and that x_ax is always larger than y_ax
#rotate X and Y is x_ax < y_ax, rotation is clockwise
if int(xy_grid['x_ax']) >= int(xy_grid['y_ax']):
  y_ax = int(xy_grid['y_ax'])
  x_ax = int(xy_grid['x_ax'])
  matrix_setting['XY_Rotate'] = False
  matrix_setting['New_Ent_Dir'] = matrix_setting['Entrance_Dir']
else:
  y_ax = int(xy_grid['x_ax'])
  x_ax = int(xy_grid['y_ax'])
  matrix_setting['XY_Rotate'] = True
  if matrix_setting['Entrance_Dir'] != 0:
    matrix_setting['New_Ent_Dir'] = matrix_setting['Entrance_Dir'] - 1
  else:
    matrix_setting['New_Ent_Dir'] = matrix_setting['Entrance_Dir'] + 3

#Flip matrix if new entrance is in the north side (1)
if matrix_setting["Entrance_Dir"] == 1:
  matrix_setting["XY_Mirror"] = True
else:
  matrix_setting["XY_Mirror"] = False

#Rotation and Mirror Function
def RotateMatrix(m_id):
  rows = len(m_id)
  cols = len(m_id[0])
  tm_id = []
  for y in range(1,cols,1):
    row = []
    for x in range(rows):
      row.append(m_id[x][-y])
    tm_id.append(row)
  return tm_id

def UnRotateMatrix(m_id):
  rows = len(m_id)
  cols = len(m_id[0])
  um_id = []
  for y in range(1,cols,1):
    row = []
    for x in range(1,rows,1):
      row.append(m_id[-x][-y])
    um_id.insert(0,row)
  return um_id

def MirrorMatrix(m_id):
  mm_id = []
  for i in range(1,len(m_id),1):
    mm_id.append(m_id[-i])
  return mm_id

#Find the general entrance location
#Create a dummy matrix to transform and check if entrance location is transformed correctly
def DummyMatrix(y_ax,x_ax,Ent_pts):
  dm_id = []
  for y in range(y_ax):
    dm_yid = []
    for x in range(x_ax):
      dm_yid.append('D')
    dm_id.append(dm_yid)
  for d in range(len(Ent_pts)):
    dm_id[Ent_pts[d][0]][Ent_pts[d][1]] = 'E'
  return dm_id

dm_id = DummyMatrix(y_ax,x_ax,Ent_pts)
if matrix_setting['XY_Rotate'] ==  True:
  dm_id = RotateMatrix(dm_id)
if matrix_setting['XY_Mirror'] ==  True:
  dm_id = MirrorMatrix(dm_id)

#From dummy matrix, find the Entrance points to determine the singular point (either in the middle, the left or right end)
def Find_EndPt(dm_id,matrix_setting,x_ax):
  Entrance_pts = []
  for y in range(len(dm_id)):
    for x in range(len(dm_id[0])):
      if dm_id[y][x] == 'E':
        Entrance_pts.append([y,x])
  if [0,0] in Entrance_pts or [0,x_ax] in Entrance_pts:
    if Entrance_pts.count([0,0]) == 1:
      matrix_setting['Ent_Pt'] = [0,0]
    elif Entrance_pts.count([0,x_ax]) == 1:
      matrix_setting['Ent_Pt'] = [0,x_ax]
  elif matrix_setting['New_Ent_Dir'] == 1 or matrix_setting['New_Ent_Dir'] == 3:
    matrix_setting['Ent_Pt'] = Entrance_pts[round(len(Entrance_pts)/2)]
  elif matrix_setting['New_Ent_Dir'] == 0 or matrix_setting['New_Ent_Dir'] == 2:
    matrix_setting['Ent_Pt'] = min(Entrance_pts)
  return None

Find_EndPt(dm_id,matrix_setting,x_ax)


"""## ASD Code

### New Layout Generation Help functions
"""

typology = "linear_adj"

### Objective to find distance between cluster & waiting rooms
# find all C:
# Self_test need specify the x and y 
def self_test(self, m_id):

    # bug
    if self[0] < 0 or self[0] >= y_ax or self[1] < 0 or self[1] >= x_ax:
        return "oob"
    else:
        return m_id[self[0]][self[1]]

def wDist(m_id, c_id):
    c_index = []
    w_index = []
    allDist = []

    for y in range(len(m_id)):
        for x in range(len(m_id[y])):
            pt = [y,x]
            if self_test(pt, m_id) == "W":
                w_index.append(pt)

    for pt in w_index:
        distance = ((c_id[1] - pt[1])**2 + (c_id[0] - pt[0])**2)**0.5
        #distance = m.dist(pt, wpt)
        allDist.append(distance)
    minDist = min(allDist)
    if minDist > 20 :
        return False
    
    elif minDist < 20 :
        return True

#Definition to find reception location if required in SOA
def GenRecLoc(m_id,y_ax,x_ax,num_dcols,matrix_setting):
  rec_size = [5,7]
  rec_loc = []
  if (x_ax - (8+(num_dcols*9))) // 3 == 0:
    for i in range(num_dcols):
      rec_loc.append([0,5+(i*9)])
  elif (x_ax - (8+(num_dcols*9))) // 3 == 1:
    for i in range(num_dcols+1):
      rec_loc.append([0,2+(i*9)])
  elif (x_ax - (8+(num_dcols*9))) // 3 == 2:
    for i in range(num_dcols+1):
      rec_loc.append([0,5+(i*9)])
  #find the distance between entrance point and possible reception points
  rec_dist = []
  for i in range(len(rec_loc)):
    ent_pt = matrix_setting['Ent_Pt']
    dist = abs(((ent_pt[0]-rec_loc[i][0])**2 + (ent_pt[1]-rec_loc[i][1])**2)**0.5)
    rec_dist.append(dist)
  GenRecLoc = rec_loc[rec_dist.index(min(rec_dist))]
  return GenRecLoc,rec_size

def GenOffice(m_id, y_ax, x_ax, num_dcols, rec_loc,col_prop):
  office_size = [6,3]
  office_loc = []
  min_col_x = col_prop['Column Index'][0][1]
  max_x_val = []
  for i in range(len(col_prop['Column Index'])):
    max_x_val.append(col_prop['Column Index'][i][1])
  max_col_x = max(max_x_val)
  if rec_loc != None:
      if rec_loc[1] <= (x_ax // 2):
          office_loc = [y_ax - office_size[0] -1 , max_col_x]
      elif rec_loc[1] > (x_ax // 2):
          office_loc = [y_ax - office_size[0] -1, min_col_x]
      office_col = 0
      for i in range(len(col_prop['Column Index'])):
          if office_loc[1] == col_prop['Column Index'][i][1]:
              office_col = i
      for y in range(office_size[0]):
        for x in range(office_size[1]):
          m_id[office_loc[0] + y][office_loc[1] + x] = "S"
  else:
      seed = ran.randint(1,2)
      if seed == 1:
        office_loc = [y_ax - office_size[0] -1 , max_col_x]
      if seed == 2:
        office_loc = [y_ax - office_size[0] -1, min_col_x]
      for i in range(len(col_prop['Column Index'])):
          if office_loc[1] == col_prop['Column Index'][i][1]:
              office_col = i
      for y in range(office_size[0]):
        for x in range(office_size[1]):
          m_id[office_loc[0] + y][office_loc[1] + x] = "S"
  return m_id,office_loc, office_size, office_col

#Definition to generate matrix and column property

#Definition to generate matrix and column property
def gen_linear_horizontal(y_ax,x_ax,matrix_setting,SOA_dict):
  col_prop = {"Column_ID" : [], "Column Index" : [], "Length" : [], "Breadth" : [], "Door" : [], "Window" : [], "Cluster Space" : [], "Room Door" : []}
  m_id = []
  rec_loc = None
  num_dcols = (x_ax-8)//9
  num_cols = 0
  col_type = (x_ax - (8+(num_dcols*9))) // 3
  #Generate m_id based on the 3 possible column config
  #Column type 1: follows a 1col + n(d_cols) + 1col configuration
  if col_type == 0:
    num_cols += (2 + (2*num_dcols))
    for i in range(y_ax-1):
      x_id = ['0','0','0']
      for j in range(num_dcols):
        x_id += ['1','1','0','0','0','2','0','0','0']
      x_id = x_id + ['1','1','0','0','0']
      m_id.append(x_id)
    last_row = []
    for j in range(8+(num_dcols*9)):
      last_row.append('2')
    m_id.append(last_row)
    col_start = 0
    for col in range(num_cols):
      if col % 2 == 0:
        col_prop['Column Index'].append([0,col_start])
        col_prop['Length'].append(y_ax -1)
        col_start += 5
      elif col % 2 != 0:
        col_prop['Column Index'].append([0,col_start])
        col_prop['Length'].append(y_ax -1)
        col_start += 4
  #Column type 2: follows a n(d_cols) configuration
  elif col_type == 1:
    num_cols += ((num_dcols+1)*2)
    for i in range(y_ax-1):
      x_id = []
      for j in range(num_dcols+1):
        x_id += ['1','1','0','0','0','2','0','0','0']
      x_id = x_id + ['1','1']
      m_id.append(x_id)
    last_row = []
    for j in range(2+((num_dcols+1)*9)):
      last_row.append('2')
    m_id.append(last_row)
    col_start = 2
    for col in range(num_cols):
      if col % 2 == 0:
        col_prop['Column Index'].append([0,col_start])
        col_prop['Length'].append(y_ax -1)
        col_start += 4
      elif col % 2 != 0:
        col_prop['Column Index'].append([0,col_start])
        col_prop['Length'].append(y_ax -1)
        col_start += 5
  #Column type 3: follows a 1col + n(d_cols) configuration
  elif col_type == 2:
    num_cols += (1 + ((num_dcols+1)*2))
    for i in range(y_ax-1):
      x_id = ['0','0','0']
      for j in range(num_dcols+1):
        x_id += ['1','1','0','0','0','2','0','0','0']
      x_id = x_id + ['1','1']
      m_id.append(x_id)
    last_row = []
    for j in range(5+((num_dcols+1)*9)):
      last_row.append('2')
    m_id.append(last_row)
    col_start = 0
    for col in range(num_cols):
      if col % 2 == 0:
        col_prop['Column Index'].append([0,col_start])
        col_prop['Length'].append(y_ax -1)
        col_start += 5
      elif col % 2 != 0:
        col_prop['Column Index'].append([0,col_start])
        col_prop['Length'].append(y_ax -1)
        col_start += 4
  #Edit the column property to account for the reception location
  if int(SOA_dict['RECEPTION']) == 1:
    Rec_pt,rec_size = GenRecLoc(m_id,y_ax,x_ax,num_dcols,matrix_setting)
    Rec_index1 = col_prop['Column Index'].index(Rec_pt)
    Rec_index2 = Rec_index1 + 1
    col_prop['Length'][Rec_index1] = y_ax - rec_size[0] - 1
    col_prop['Length'][Rec_index2] = y_ax - rec_size[0] - 1
    col_prop['Column Index'][Rec_index1] = [Rec_pt[0]+5,Rec_pt[1]]
    col_prop['Column Index'][Rec_index2] = [Rec_pt[0]+5,Rec_pt[1]+4]
    rec_loc = Rec_pt
    for y in range(rec_size[0]):
      for x in range(rec_size[1]):
        m_id[Rec_pt[0]+y][Rec_pt[1]+x] = "R"
  else:
    pass
  
  
  if int(SOA_dict['STAFF OFFICE']) == 1:
    m_id, office_loc, office_size, office_col = GenOffice(m_id, y_ax, x_ax, num_dcols, rec_loc,col_prop)
    col_prop['Length'][office_col] = col_prop['Length'][office_col] - office_size[0]
        
  else:
    pass

  for col in range(num_cols):
      col_prop['Column_ID'].append(col+1)
      col_prop['Breadth'].append(3)
  
  #print (col_prop['Column Index'])
  #Add connecting corridor if y_ax is more than 20 (24m)
  if y_ax > 20:
    mid_y = round(y_ax/2)
    new_col_index = []
    #change length of columns in between column index 0 and -1
    for i in range(len(col_prop['Column_ID'])):
      if i == 0  or i == (len(col_prop['Column_ID'])-1):
        pass
      else:
        col_prop['Length'][i] = col_prop['Length'][i] - mid_y
        new_col_index.append(col_prop['Column Index'][i])
    #change mid_y and mid_y+1 into corridor
    #find x value of the flanking rooms
    end_x1 = col_prop['Column Index'][0][1]+2
    end_x2 = col_prop['Column Index'][-1][1]
    #print (end_x1,end_x2)
    for j in range(2):
      for x in range(x_ax):
        if x > end_x1 and x < end_x2:
          m_id[mid_y+j][x] = '1'
    #create new columns
    #print (len(new_col_index))
    for k in range(len(new_col_index)):
      col_prop["Column Index"].append([mid_y+2,new_col_index[k][1]])
      col_prop['Column_ID'].append(col_prop['Column_ID'][len(new_col_index)]+k+2) 
      col_prop['Length'].append(y_ax - mid_y - 3)
      col_prop['Breadth'].append(3)

  #find room door direction of each column and sort be column number
  for col in range(len(col_prop['Column_ID'])):
    test_id = col_prop['Column Index'][col]
    if test_id[1] + 3 < len(m_id[0]):
        leftside = self_test([test_id[0],test_id[1]-1],m_id)
        rightside = self_test([test_id[0],test_id[1]+3],m_id)
        if leftside == '1':
            col_prop["Room Door"].append(0)
        elif rightside == '1':
            col_prop["Room Door"].append(1)
    elif test_id[1] + 3 >= len(m_id[0]):
        col_prop["Room Door"].append(0)

  return m_id,col_prop,rec_loc, office_loc, office_col

m_id,col_prop,rec_loc, office_loc, office_col = gen_linear_horizontal(y_ax,x_ax,matrix_setting,SOA_dict)
totalElem = 0
for i in range(len(col_prop['Length'])):
  totalElem += (col_prop['Length'][i]//2)

#add staff corridor walls: test all '2' in m_id if it is adjacent a 1
SC_wall = []
SC_id = []
for y in range(len(m_id)):
  for x in range(len(m_id[y])):
    if m_id[y][x] == '2':
      SC_id.append([y,x])
    else:
      pass 

#test top and bottom of corridor
for i in SC_id:
  if self_test([i[0]+1,i[1]],m_id) != 'oob' and self_test([i[0]-1,i[1]],m_id) != 'oob':
    
    if m_id[i[0]+1][i[1]] == '1':
      SC_pt1 = [i[0]+1,i[1]]
      SC_pt2 = [i[0]+1,i[1]+1]
      SC_wall.append([SC_pt1,SC_pt2])
    elif m_id[i[0]-1][i[1]] == '1':
      SC_pt1 = [i[0],i[1]]
      SC_pt2 = [i[0],i[1]+1]
      SC_wall.append([SC_pt1,SC_pt2])
  else:

      if i[0] == 0:
        SC_pt1 = [i[0],i[1]]
        SC_pt2 = [i[0],i[1]+1]
        SC_wall.append([SC_pt1,SC_pt2])
        
      else:
        if m_id[i[0]-1][i[1]] == '1':
          SC_pt1 = [i[0],i[1]]
          SC_pt2 = [i[0],i[1]+1]
          SC_wall.append([SC_pt1,SC_pt2])
        else:
            pass



#Definition to transform cluster
def trans_cl(cluster_id,trans_type):
  if trans_type == 0:
    return cluster_id
  elif trans_type == 1:
    cluster_id.reverse()
    return cluster_id
  elif trans_type == 2:
    for sublist in cluster_id:
      sublist.reverse()
    return cluster_id
  elif trans_type == 3:
    cluster_id.reverse()
    for sublist in cluster_id:
      sublist.reverse()
    return cluster_id
  else:
    print("Input is invalid")
    return cluster_id

#Generate Clusters
def RL(t_id, m_id, transform, y_ax, x_ax):
  cluster_id = [['R','R','R','R','R','R','R'],
                ['R','R','R','R','R','R','R'],
                ['R','R','R','R','R','R','R'],
                ['W','W','W','W','W','W','W'],
                ['W','W','W','W','W','W','W']]           

  cl_prop = {"Cluster Name": "RL", "Cluster Name(H)" : "RH", "Corridor Area": 0, "Waiting Room": 1, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 1, 
             "Waiting Area": 20.16, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 5, "Breadth" : 7, "Column Index" : 99, "Valid" : 0}
  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  if bounds[0] > y_ax or bounds[1] > x_ax:
      cl_prop = {"Cluster Name": "RL", "Cluster Name(H)" : "RH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception":0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 4, "Breadth" : 3, "Column Index" : 99, "Valid" : 0}
      return cl_prop, m_id, t_id
  else: 
      cl_prop['Valid'] = 1
      new_t_id = [t_id[0] + cl_prop["Length"], t_id[1]]
      t_cluster_id = trans_cl(cluster_id, transform)
      for y in range(cl_prop['Length']):
        for x in range(cl_prop['Breadth']):
          m_id[t_id[0] + y][t_id[1] + x] = t_cluster_id[y][x]
      return cl_prop, m_id, new_t_id

#New Toilet Cluster 1
def TL(t_id, transform, m_id, y_ax, x_ax, column_number):
  cluster_id = [['T','T','T'],
                ['T','T','T']]

  cl_prop = {"Cluster Name": "TL", "Cluster Name(H)" : "TL", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 1, "Staff Room": 0,  "Turns" :0, "Length" : 2, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  mid_cor = 0
  for y in range(cl_prop['Length']):   
    if t_id[0]+y < y_ax:
      if self_test([t_id[0]+y,t_id[1]],m_id) != '0':
        mid_cor += 1
  if bounds[0] > y_ax or bounds[1] > x_ax:
      cl_prop = {"Cluster Name": "TL", "Cluster Name(H)" : "TL", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 2, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  elif mid_cor > 0:
      cl_prop = {"Cluster Name": "TL", "Cluster Name(H)" : "TL", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 1, "Staff Room": 0,  "Turns" :0, "Length" : 2, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  else:    
      cl_prop['Valid'] = 1
      new_t_id = [t_id[0] + cl_prop["Length"], t_id[1]]
      t_cluster_id = trans_cl(cluster_id, transform)
      for y in range(cl_prop['Length']):
        for x in range(cl_prop['Breadth']):
          m_id[t_id[0] + y][t_id[1] + x] = t_cluster_id[y][x]
      return cl_prop, m_id, new_t_id


#New Consultation Room Cluster 1
def CE1L(t_id, transform, m_id, y_ax, x_ax, column_number):
  cluster_id = [['C','C','C'],
                ['C','C','C'],
                ['C','C','C']]

  cl_prop = {"Cluster Name": "CE1L", "Cluster Name(H)" : "CE1H", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 1, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 3, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}

  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  mid_cor = 0
  for y in range(cl_prop['Length']):   
    if t_id[0]+y < y_ax:
      if self_test([t_id[0]+y,t_id[1]],m_id) != '0':
        mid_cor += 1
  if bounds[0] > y_ax or bounds[1] > x_ax:
      cl_prop = {"Cluster Name": "CE1L", "Cluster Name(H)" : "CE1H", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
              "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 3, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  elif mid_cor > 0:
      cl_prop = {"Cluster Name": "CE1L", "Cluster Name(H)" : "CE1H", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
              "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 3, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  else:  
      cl_prop['Valid'] = 1
      new_t_id = [t_id[0] + cl_prop["Length"], t_id[1]]
      t_cluster_id = trans_cl(cluster_id, transform)
      for y in range(cl_prop['Length']):
        for x in range(cl_prop['Breadth']):
          m_id[t_id[0] + y][t_id[1] + x] = t_cluster_id[y][x]
      return cl_prop, m_id, new_t_id


#New Consultation Room Cluster 1
def CE2L(t_id, transform, m_id, y_ax, x_ax, column_number):
  cluster_id = [['C','C','C'],
                ['C','C','C'],
                ['C','C','C'],
                ['C','C','C']]

  cl_prop = {"Cluster Name": "CE2L", "Cluster Name(H)" : "CE2H", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 1, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 4, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}

  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  mid_cor = 0
  for y in range(cl_prop['Length']):   
    if t_id[0]+y < y_ax:
      if self_test([t_id[0]+y,t_id[1]],m_id) != '0':
        mid_cor += 1
  if bounds[0] > y_ax or bounds[1] > x_ax:
      cl_prop = {"Cluster Name": "CE2L", "Cluster Name(H)" : "CE2H", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
              "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 4, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  elif mid_cor > 0:
      cl_prop = {"Cluster Name": "CE2L", "Cluster Name(H)" : "CE2H", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
              "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 4, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  else:  
      cl_prop['Valid'] = 1
      new_t_id = [t_id[0] + cl_prop["Length"], t_id[1]]
      t_cluster_id = trans_cl(cluster_id, transform)
      for y in range(cl_prop['Length']):
        for x in range(cl_prop['Breadth']):
          m_id[t_id[0] + y][t_id[1] + x] = t_cluster_id[y][x]
      return cl_prop, m_id, new_t_id

#Staff Room 3x5
def SL(t_id, m_id, transform, y_ax, x_ax, column_number):
  cluster_id = [['S','S','S'],
                ['S','S','S'],
                ['S','S','S'],
                ['S','S','S'],
                ['S','S','S'],
                ['S','S','S']]

  cl_prop = {"Cluster Name": "SL", "Cluster Name(H)" : "SH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': 0, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 1, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 1,  "Turns" :0, "Length" : 6, "Breadth" : 3, "Column Index" : 99, "Valid" : 1}
  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  return cl_prop, m_id, t_id


#Interview Rooms
def IL(t_id, transform, m_id, y_ax, x_ax, column_number):
  cluster_id = [['I','I','I'],
                ['I','I','I'],
                ['I','I','I']]

  cl_prop = {"Cluster Name": "IL", "Cluster Name(H)" : "IH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': 0, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 1, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 3, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  mid_cor = 0
  for y in range(cl_prop['Length']):   
    if t_id[0]+y < y_ax:
      if self_test([t_id[0]+y,t_id[1]],m_id) != '0':
        mid_cor += 1
  if bounds[0] > y_ax or bounds[1] > x_ax:
      return cl_prop, m_id, t_id

  elif mid_cor > 0:
      cl_prop = {"Cluster Name": "IL", "Cluster Name(H)" : "IH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': 0, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 1, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 3, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  else:
      cl_prop['Valid'] = 1
      new_t_id = [t_id[0] + cl_prop["Length"], t_id[1]]
      t_cluster_id = trans_cl(cluster_id, transform)
      for y in range(cl_prop['Length']):
        for x in range(cl_prop['Breadth']):
          m_id[t_id[0] + y][t_id[1] + x] = t_cluster_id[y][x]
      return cl_prop, m_id, new_t_id



#Clean Utility Room
def CDL(t_id, transform, m_id, y_ax, x_ax, column_number):
  cluster_id = [['CU','CU','CU'],
                ['CU','CU','CU'],
                ['CU','CU','CU'],
                ['CU','CU','CU'],
                ['DU','DU','DU'],
                ['DU','DU','DU'],
                ['DU','DU','DU'],
                ['DU','DU','DU']]

  cl_prop = {"Cluster Name": "CDL", "Cluster Name(H)" : "CDH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': 0, "C/E Rm": 0, "Clean Utility": 1, "Dirty Utility": 1, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 8, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  mid_cor = 0
  for y in range(cl_prop['Length']):   
    if t_id[0]+y < y_ax:
      if self_test([t_id[0]+y,t_id[1]],m_id) != '0':
        mid_cor += 1
  if bounds[0] > y_ax or bounds[1] > x_ax:
      return cl_prop, m_id, t_id    
  elif mid_cor > 0:
      cl_prop =  {"Cluster Name": "CDL", "Cluster Name(H)" : "CDH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': 0, "C/E Rm": 0, "Clean Utility": 1, "Dirty Utility": 1, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 8, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  else:
      cl_prop['Valid'] = 1
      new_t_id = [t_id[0] + cl_prop["Length"], t_id[1]]
      t_cluster_id = trans_cl(cluster_id, transform)
      for y in range(cl_prop['Length']):
        for x in range(cl_prop['Breadth']):
          m_id[t_id[0] + y][t_id[1] + x] = t_cluster_id[y][x]
      return cl_prop, m_id, new_t_id


#New Waiting Room Cluster 1
def WSL(t_id, transform, m_id, y_ax, x_ax, column_number):
  cluster_id = [['W','W','W'],
                ['W','W','W'],
                ['W','W','W']]

  cl_prop = {"Cluster Name": "WSL", "Cluster Name(H)" : "WSH", "Corridor Area": 0, "Waiting Room": 1, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 12.96, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 3, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  mid_cor = 0
  for y in range(cl_prop['Length']):   
    if t_id[0]+y < y_ax:
      if self_test([t_id[0]+y,t_id[1]],m_id) != '0':
        mid_cor += 1
  if bounds[0] > y_ax or bounds[1] > x_ax:
      cl_prop = {"Cluster Name": "WSL", "Cluster Name(H)" : "WSH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 6, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 3, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop,m_id, t_id
  elif mid_cor > 0:
      cl_prop = {"Cluster Name": "WSL", "Cluster Name(H)" : "WSH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 6, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 3, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  else:

      cl_prop['Valid'] = 1
      new_t_id = [t_id[0] + cl_prop["Length"], t_id[1]]
      t_cluster_id = trans_cl(cluster_id, transform)
      for y in range(cl_prop['Length']):
        for x in range(cl_prop['Breadth']):
          m_id[t_id[0] + y][t_id[1] + x] = t_cluster_id[y][x]
      return cl_prop, m_id, new_t_id

#New Waiting Room Cluster 2
def WBL(t_id, transform, m_id, y_ax, x_ax, column_number):
  cluster_id = [['W','W','W'],
                ['W','W','W'],
                ['W','W','W'],
                ['W','W','W'],
                ['W','W','W'],
                ['W','W','W']]

  cl_prop = {"Cluster Name": "WBL", "Cluster Name(H)" : "WBH", "Corridor Area": 0, "Waiting Room": 1, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 25.92, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 6, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  mid_cor = 0
  for y in range(cl_prop['Length']):   
    if t_id[0]+y < y_ax:
      if self_test([t_id[0]+y,t_id[1]],m_id) != '0':
        mid_cor += 1
  if bounds[0] > y_ax or bounds[1] == x_ax:
      cl_prop = {"Cluster Name": "WBL", "Cluster Name(H)" : "WBH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 6, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop,m_id, t_id
  elif mid_cor > 0:
      cl_prop = {"Cluster Name": "WBL", "Cluster Name(H)" : "WBH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 6, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop,m_id, t_id
  else:

      cl_prop['Valid'] = 1
      new_t_id = [t_id[0] + cl_prop["Length"], t_id[1]]
      t_cluster_id = trans_cl(cluster_id, transform)
      for y in range(cl_prop['Length']):
        for x in range(cl_prop['Breadth']):
          m_id[t_id[0] + y][t_id[1] + x] = t_cluster_id[y][x]
      return cl_prop, m_id, new_t_id


#New Waiting Room Cluster 3
def WSXL(t_id, transform, m_id, y_ax, x_ax, column_number):
  cluster_id = [['W','W','W'],
                ['W','W','W']]

  cl_prop = {"Cluster Name": "WSXL", "Cluster Name(H)" : "WSXH", "Corridor Area": 0, "Waiting Room": 1, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 8.6, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 2, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
  bounds = [t_id[0] + cl_prop["Length"], t_id[1] + cl_prop["Breadth"]]
  mid_cor = 0
  for y in range(cl_prop['Length']):   
    if t_id[0]+y < y_ax:
      if self_test([t_id[0]+y,t_id[1]],m_id) != '0':
        mid_cor += 1
  if bounds[0] > y_ax or bounds[1] > x_ax:
      cl_prop = {"Cluster Name": "WSL", "Cluster Name(H)" : "WSH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 6, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 2, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop,m_id, t_id
  elif mid_cor > 0:
      cl_prop = {"Cluster Name": "WSL", "Cluster Name(H)" : "WSH", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 6, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 2, "Breadth" : 3, "Column Index" : column_number, "Valid" : 0}
      return cl_prop, m_id, t_id
  else:

      cl_prop['Valid'] = 1
      new_t_id = [t_id[0] + cl_prop["Length"], t_id[1]]
      t_cluster_id = trans_cl(cluster_id, transform)
      for y in range(cl_prop['Length']):
        for x in range(cl_prop['Breadth']):
          m_id[t_id[0] + y][t_id[1] + x] = t_cluster_id[y][x]
      return cl_prop, m_id, new_t_id


#Empty Cluster
def Empty_Cluster(t_id, transform, m_id, column_number):

  cl_prop = {"Cluster Name": "Empty_Cluster", "Cluster Name(H)" : "Empty", "Corridor Area": 0, "Waiting Room": 0, 'Waiting Room Dist': -1, "C/E Rm": 0, "Clean Utility": 0, "Dirty Utility": 0, "Interview": 0, "Office": 0, "Reception": 0, 
             "Waiting Area": 0, "Toilet": 0, "Staff Room": 0,  "Turns" :0, "Length" : 0, "Breadth" : 0, "Column Index" : column_number, "Valid" : 0}

  return cl_prop, m_id, t_id

#Call for different clusters to be placed in the columns
#GenRC is used to generate the reception in grid with the door property
#Takes in the RC number, and the arguements of the cluster definition
def genRC(RC_num,transform, t_id, m_id,y_ax, x_ax, column_number):
    if RC_num == 1:
        return RL(t_id ,m_id, transform, y_ax, x_ax)
    elif RC_num == 2:
        return SL(t_id ,m_id, transform, y_ax, x_ax, column_number)
  #elif RC_num == 2:
    #return R3X5(t_id, m_id, transform, y_ax, x_ax, col_prop, column_number)
    else:
        print("Invalid input to generate reception cluster")
        return None

def genAll(num,transform, t_id,m_id, y_ax, x_ax, column_number):
  if   num == 1:
    return CE1L(t_id, transform, m_id, y_ax, x_ax, column_number)
  elif num == 2:
    return CE2L(t_id, transform, m_id, y_ax, x_ax, column_number)
  elif num == 3:
    return IL(t_id, transform, m_id, y_ax, x_ax, column_number)
  elif num == 4:
    return WSXL(t_id, transform, m_id, y_ax, x_ax, column_number)
  elif num == 5:
    return WSL(t_id, transform, m_id, y_ax, x_ax, column_number)
  elif num == 6:
    return WBL(t_id, transform, m_id, y_ax, x_ax, column_number)
  #elif num == 6:
  #  return TL(t_id, transform, m_id, y_ax, x_ax, column_number)
  elif num == 7:
    return Empty_Cluster(t_id, transform, m_id, column_number)
  else:
    print("Invalid input to generate consultation room cluster.")
    return None

def gen_mid(m_id,dSpace,cluster_prop,col_prop,keylist, y_ax, x_ax, testCR, testCRTR, testR, testRT,rec_loc, office_loc, office_col):
    if rec_loc != None:
        cl_prop, m_id, t_id = genRC(1, 0, rec_loc, m_id, y_ax, x_ax, 99)
        for j in range(len(keylist)-1):
                cluster_prop[str(keylist[j])].append(cl_prop[str(keylist[j])])
        cluster_prop['Cluster ID'].append(rec_loc)
    else:
        pass
    if office_loc != None:
        if office_loc[1] > x_ax // 2:
            cl_prop, m_id, t_id = genRC(2, 0, office_loc, m_id, y_ax, x_ax, office_col)
            for j in range(len(keylist)-1):
                    cluster_prop[str(keylist[j])].append(cl_prop[str(keylist[j])])
        elif office_loc[1] < x_ax // 2:
            cl_prop, m_id, t_id = genRC(2, 0, office_loc, m_id, y_ax, x_ax, 1)
            for j in range(len(keylist)-1):
                    cluster_prop[str(keylist[j])].append(cl_prop[str(keylist[j])])
        cluster_prop['Cluster ID'].append(office_loc)
    else:
        pass

    col_pt = col_prop['Column Index']
    counter = 0
    for i in range(len(col_pt)):

        column_number = i + 1
        t_id = col_pt[i]
        #print (dSpace)
        #print (testCR)
        #print (testCRTR)
        #print (t_id) 
        #print (column_number)
        for d in range(dSpace[i]):
            prev_t_id = t_id
            #print (counter)
            cl_prop, m_id, t_id = genAll(testCR[counter], testCRTR[counter], t_id, m_id, y_ax, x_ax, column_number)
            for j in range(len(keylist)-1):
                cluster_prop[str(keylist[j])].append(cl_prop[str(keylist[j])])
            cluster_prop['Cluster ID'].append(prev_t_id)
            counter += 1
            #print (column_number)
            #print (prev_t_id)

    return cluster_prop, m_id

"""### ASD Layout main code"""

# TYPOLOGY FUNCTIONS
## Horizontal Linear Typology
def Layout(typology, y_ax, x_ax, testCR, testCRTR, testR, testRT, matrix_setting, SOA_dict):
  #y_ax =  number of rows, x_ax =  number of columns
  #Input from SOA: grid size
  #Instantiate the matrix
  m_id,col_prop,rec_loc, office_loc, office_col = gen_linear_horizontal(y_ax,x_ax,matrix_setting,SOA_dict)
  col_pt = col_prop["Column Index"]
  testR = 1
  testRT = 0
  dSpace = []
  num_cols = len(col_prop['Column_ID'])
  totalElem = 0
  for col in range(len(col_prop['Column_ID'])):
    totalElem += (col_prop['Length'][col]) // 2
    dSpace.append((col_prop['Length'][col]) // 2)

  
  cluster_prop = {"Cluster Name":[], "Corridor Area":[], "Waiting Room":[], 'Waiting Room Dist':[], "C/E Rm":[], "Clean Utility":[], "Dirty Utility":[], "Interview":[], "Office":[], "Reception":[], "Waiting Area":[], "Toilet":[], "Staff Room":[],  "Turns" : [], 
                    "Column Index" : [], "Valid" :[], "Length" :[], "Breadth" : [], "Cluster ID":[]}
  keylist = list(cluster_prop)
  ncluster_prop, nm_id = gen_mid(m_id,dSpace,cluster_prop,col_prop,keylist, y_ax, x_ax, testCR, testCRTR, testR, testRT,rec_loc, office_loc, office_col)

  return [nm_id, ncluster_prop,col_prop, totalElem]

testCR = []
testCRTR = []
i = 0
while i < totalElem:
    testCRTR.append(0)
    testCR.append(ran.randint(1,6))
    #testCR.append(1)
    i += 1
#testCR = [1, 1, 9, 1, 5, 5, 1, 9, 5,
          #1, 5, 1 ,2, 3, 5, 5, 5, 5, 1, 
          #5, 4, 1, 1, 1, 1, 9, 4 ,5, 5]
#testCRTR = [0, 0, 0, 0, 0, 0, 0, 0 ,0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0 ,0, 0, 0, 0, 0, 0]
#print(len(testCR))
testR=1
testRT=0
test_result = Layout(typology, y_ax, x_ax, testCR, testCRTR, testR, testRT,matrix_setting,SOA_dict)
#totalElem = test_result[3]
m_id = test_result[0]


"""## ESD Code

Parameters :
- m_id：The list of the layout matrix
- cluster_prop : property for cluster info
- grid_prop: property info for grid matrix

### Help functions
"""

# Visualize function 

def visualization_matrix(m_id):
  fig, ax = plt.subplots()


  x_min,x_max=0,len(m_id[0])
  y_min,y_max=0,len(m_id)
  for i in range(x_max):
      for j in range(y_max):
          c = m_id[j][i]
          ax.text(i+0.5, j+0.5, c, va='center', ha='center')


  ax.set_xlim(x_min, x_max)
  ax.set_ylim(y_min, y_max)
  ax.set_xticks(np.arange(x_max))
  ax.set_yticks(np.arange(y_max))
  ax.grid()
visualization_matrix(m_id)

# function to check the adj 
def adj_test(self, m_id, subject):
    results = []
    #Test left
    if self[1] - 1 < 0:
        results.append("oob")
    elif m_id[self[0]][self[1] - 1] == subject:
         #print m_id[self[0]][self[1] - 1]
         results.append(True)
    else:
        results.append(False)
    #Test right
    if self[1] + 1 >= 8:
        results.append("oob")
    elif m_id[self[0]][self[1]+1] == subject:
        results.append(True)
    else:
        results.append(False)
    #Test up
    if self[0] + 1 >= 7:
        results.append("oob")
    elif m_id[self[0]+1][self[1]] == subject:
        results.append(True)
    else:
        results.append(False)
    #Test down
    if self[0] - 1 < 0:
        results.append("oob")
    elif m_id[self[0]-1][self[1]] == subject:
        results.append(True)
    else:
        results.append(False)
    return True in results

# find all C:
# Self_test need specify the x and y 
def self_test(self, m_id):

    # bug
    if self[0] < 0 or self[0] >= y_ax or self[1] < 0 or self[1] >= x_ax:
        return "oob"
    else:
        return m_id[self[0]][self[1]]

def isValid(x, y, grid, visited):
    if ((x >= 0 and y >= 0) and
        (x < len(grid) and y < len(grid[0])) and
            (grid[x][y] != '0') and (visited[x][y] == False)):
        return True
    return False

"""#### Main objective 2 functions

- If there is no path existed then return the distance with the MAX_DISTANCE values
"""

# Circulation: corridor + waiting room area 
def obj1(m_id):
    c_area =0
    w_area =0
    for y in range(len(m_id)):
        for x in range(len(m_id[y])):
            pt = [y,x]
            if self_test(pt, m_id) == "C":
                c_area+=1
            if self_test(pt, m_id) == "W":
                w_area+=1
    circulation_area = c_area+w_area
    return circulation_area
#obj1(cluster_prop)

## Minimize waiting area
## There is actually a similar obj function - refer to obj5
## In this obj we will count how many w in m_id instead 
def obj2(m_id):
  w_area =0
  for y in range(len(m_id)):
      for x in range(len(m_id[y])):
          pt = [y,x]
          if self_test(pt, m_id) == "W":
              w_area+=1
  return w_area

"""### Objective 3
- Find average distance of all consultation rooms to all waiting rooms
"""

def obj3(m_id):
    c_index = []
    w_index = []
    
    for y in range(len(m_id)):
        for x in range(len(m_id[y])):
            pt = [y,x]
            if self_test(pt, m_id) == "C":
                c_index.append(pt)
            if self_test(pt, m_id) == "W":
                w_index.append(pt)


    #for pt in m_id:
    #    if pt == "C":
    #        c_index.append(pt)
    #    elif pt == "W":
    #        w_index.append(pt)
    #    else:
    #        pass
    if len(c_index) == 0 or len(w_index) == 0:
        return 20
    else:
        allDist = []
        for pt in c_index:
            for wpt in w_index:
                
                distance = ((pt[1] - wpt[1])**2 + (pt[0] - wpt[0])**2)**0.5
                #distance = m.dist(pt, wpt)
                allDist.append(distance)
        avgDist = np.mean(allDist)
        return avgDist

def obj4(cluster_prop):
  waitAreas= sum(cluster_prop['Waiting Room'])
  return waitAreas

### Objective 5

#Number of waiting areas


def obj5(m_id):
    s_index = []
    r_index = []
    
    for y in range(len(m_id)):
        for x in range(len(m_id[y])):
            pt = [y,x]
            if self_test(pt, m_id) == "S":
                s_index.append(pt)
            if self_test(pt, m_id) == "R":
                r_index.append(pt)


    if len(s_index) == 0 or len(r_index) == 0:
        return x_ax
        
    else:
        allDist = []
        for pt in s_index:
            for wpt in r_index:
                
                distance = ((pt[1] - wpt[1])**2 + (pt[0] - wpt[0])**2)**0.5
                #distance = m.dist(pt, wpt)
                allDist.append(distance)
        avgDist = np.mean(allDist)
        
        return -avgDist

#Minimizing the average distance from utility rooms to the staff office.
def obj6(m_id):
    CD_index = []
    s_index = []
    
    for y in range(len(m_id)):
        for x in range(len(m_id[y])):
            pt = [y,x]
            if self_test(pt, m_id) == "CU" or "DU":
                CD_index.append(pt)
            if self_test(pt, m_id) == "S":
                s_index.append(pt)


    if len(s_index) == 0 or len(CD_index) == 0:
        return x_ax
    else:
        allDist = []
        for pt in s_index:
            for wpt in CD_index:
                
                distance = ((pt[1] - wpt[1])**2 + (pt[0] - wpt[0])**2)**0.5
                #distance = m.dist(pt, wpt)
                allDist.append(distance)
        avgDist = np.mean(allDist)
        return avgDist


#Maximise number of consultation rooms 
def obj7(cluster_prop):
  CE_rooms= sum(cluster_prop['C/E Rm'])
  return -CE_rooms


"""### Constraint 1
- User input from Excel?
"""

#Constraint 1
#Required number of consultation rooms should meet user input value or greater

def constraint1(cluster_prop):
  sum_c=sum(cluster_prop['C/E Rm'])
  return sum_c

"""### Constraint 2
- Connections of corridor
- Use grid_prop connectivity  
"""

#Constraint 2
#Each cluster should be connected to atleast one more cluster 

def constraint2(cluster_prop):
  sumInterview=sum(cluster_prop["Interview"])
  return sumInterview

"""###Constraint 3

- Interview space should be limited to 1
"""

###Constraint 3

#Interview space should be limited to 1

def constraint3(cluster_prop):
  sumOffice=sum(cluster_prop["Office"])
  return sumOffice
#constraint3(cluster_prop)

"""### Constraint 4

- Office and staff area clusters should be limited to atleast 1?
- By making limiting interview space and office space we will automatically limit staff space as well. As staff space is included in both cluster types
"""

### Constraint 4

#Office and staff area clusters should be limited to atleast 1?
#By making limiting interview space and office space we will automatically limit staff space as well. As staff space is included in both cluster types

def constraint4(cluster_prop):
  sumUtility=sum(cluster_prop["Dirty Utility"])
  return sumUtility

"""### Constraint 5

- Clean and Dirty Utility
"""

### Constraint 5

#Clean and Dirty Utility


def constraint5(cluster_prop):
  sumToilet=sum(cluster_prop["Toilet"])
  return sumToilet

### Constraint 7
#No empty space shown finally

#PS:
#the total width of all non-empty room in one column should sum to the height of the layout
#1. 9 is empty room
#2. R is alway place at col2 first row 

def constraint7(cluster_prop, col_prop):
  columnlist=[]
  for i in col_prop["Column_ID"]:
    clist=[]
    for j in range(0,len(cluster_prop["Column Index"])):
      if cluster_prop["Column Index"][j]==i:
        clist.append(cluster_prop["Length"][j])
    columnlist.append(sum(clist))
  
  return columnlist

def constraint8(cluster_prop):
  sumCUtility=sum(cluster_prop["Clean Utility"])
  return sumCUtility

def constraint7_2(cluster_prop, col_prop):
  columnlist=[]
  for i in col_prop["Column_ID"]:
    clist=[]
    for j in range(0,len(cluster_prop["Column Index"])):
      if cluster_prop["Column Index"][j]==i and cluster_prop["Valid"][j]==1:
        clist.append(cluster_prop["Length"][j])
    columnlist.append(sum(clist))

  return columnlist

""".## NSGA Algorithm 

- Objective functions x2:
  1. obj3
  2. obj5
  
"""

# Selection function for objs 
def obj_switch(n,m_id, cluster_prop ,col_prop):
  if n==1:
    return obj2(m_id)
  if n==2:
    return obj3(m_id)
  if n==3:
    return obj7(cluster_prop)

def con_switch(n,m_id, cluster_prop ,col_prop,SOA_dict):
  if n==1:
    num_consult_rooms=SOA_dict["CONSULTATION/EXAMINATION ROOM"]
    G1= num_consult_rooms -constraint1(cluster_prop)
    return G1
  if n==2:
    num_interview_rooms=SOA_dict["INTERVIEW ROOM"]
    G2 = num_interview_rooms-constraint2(cluster_prop)
    return G2
  if n==3:
    return abs(1-constraint3(cluster_prop))
  if n==4:
    return abs(1-constraint4(cluster_prop))
  #if n==5:
  #  G5 = (SOA_dict["TOILET"]-constraint5(cluster_prop))
  #  return G5
  if n==6:
    G6 = (SOA_dict["WAITING"] - obj4(cluster_prop))
    return G6
  if n==7:
    #16,20,20,15,15,20
    #G1 = abs((16-constraint7(cluster_prop,col_prop)[0]))
    #G2 = abs((20-constraint7(cluster_prop,col_prop)[1]))
    #G3 = abs((20-constraint7(cluster_prop,col_prop)[2]))
    #G4 = abs((15-constraint7(cluster_prop,col_prop)[3]))
    #G5 = abs((15-constraint7(cluster_prop,col_prop)[4]))
    #G6 = abs((20-constraint7(cluster_prop,col_prop)[5]))

    #GI_LIST=[G1,G2,G3,G4,G5,G6]
    #return GI_LIST

    GI_LIST=[]
    for i in range(len(col_prop["Column_ID"])):
      Gi = abs(col_prop['Length'][i]-constraint7(cluster_prop,col_prop)[i]) #need to decide between constraint7 and constraint7_2 
      GI_LIST.append(Gi)
    return GI_LIST

  if n==8:
    G8=abs(1-constraint8(cluster_prop))
    return G8
  #if n==9:
  #  G9 = (constraint5(cluster_prop)-(SOA_dict["TOILET"]+2))
  #  return G9
  if n==10:
    G10 = (obj4(cluster_prop)-(SOA_dict["WAITING"]+3))
    return G10
  if n==11:
    num_interview_rooms=SOA_dict["INTERVIEW ROOM"]
    G11 = constraint2(cluster_prop)-(num_interview_rooms+1)
    return G11

#Initial population creation 

#No. of columns - 10
#Column Lengths - [15, 10, 10, 5, 5, 20, 8, 8, 8, 8]

#def InitialPop(col_prop):
#  num_toilets=SOA_dict['TOILET']
#  num_consultation=SOA_dict['CONSULTATION/EXAMINATION ROOM']
#  num_waiting=SOA_dict['WAITING']
#  num_interview=SOA_dict['INTERVIEW ROOM']

#  Lengths = {'1':3,'2':4,'3':2,'4':8,'5':3,'6':6,'7':4,'8':5,'9':0}

#  for i in col_prop['Length']:
#    for j in range(0,col_prop['Legnth'][i]):
#      column_fit=0
#      if column_fit!=col_prop['Length'][i]:

"""### Finetune the parameters for NSGA
Find thebest paramter to train a model that can converge fast and output nicely 
"""

#for n_termination in range(100,1100,100):
  #""""
  #pop_size=50
  #n_offsprings=50
  #under conditon with genration n =200
  #""""
  # run the nsga
  # record res.CV in a list 
  # record n_termination 
# PLOT using pyplot 
#plt.(....)

#SPECIFYING OBJECTIVES AND PARAMETERS
objective_list = []
obj_counter = 0
for objective in TBO_Obj:
    obj_counter += 1
    if TBO_Obj[objective] == "Y":
        objective_list.append(obj_counter)
    else:
        pass

con_number = len(col_prop['Column Index'])


class MyProblem_Customized(ElementwiseProblem):
    def __init__(self,objs_list,cons_list):
        super().__init__(n_var=(totalElem), n_obj=2, n_constr=4+con_number,
                         xl=np.repeat([1],[totalElem]), 
                         xu=np.repeat([7],[totalElem]), type_var=int)
        self.objs_list= objective_list
        self.cons_list=cons_list
        

    def _evaluate(self, x, out, *args, **kwargs):

      # clusster index for first 30 and later 30 is rotation
      testCR=x[0:totalElem]
      testCRTR=np.repeat([0],[totalElem])
      testR=1 #x[totalElem-1]
      testRT=0 #x[(totalElem*2) - 1]
      #print(len(testCR))

      # Layout generation with the generated cluster index lists :
      m_id, cluster_prop ,col_prop, uselessElem =Layout(typology, y_ax, x_ax, testCR.tolist(),testCRTR.tolist(),testR,testRT,matrix_setting,SOA_dict)

      # Objs list selected :
      F_list=[]
      for oi in self.objs_list:
        F_list.append(obj_switch(oi,m_id, cluster_prop ,col_prop))

      # Cons list 
      C_list=[]
      for ci in self.cons_list:
        c_out=con_switch(ci,m_id, cluster_prop ,col_prop,SOA_dict)
        if isinstance(c_out, list):
          C_list.extend(c_out)
        else:
          C_list.append(c_out)

      # Objectives 
      out["F"] = F_list
      # Constraints
      out["G"] = C_list

print ("Lets find some fun facts!")
np.random.seed(10)
#ran.seed(1)
pop_list=[]
for i in range(0,2000):
    pop = (np.hstack((np.repeat([ran.randint(1,2)],[int(SOA_dict['CONSULTATION/EXAMINATION ROOM'])]), 
               np.repeat([ran.randint(4,6)],[int(SOA_dict['WAITING'])]),np.repeat([3],[int(SOA_dict['INTERVIEW ROOM'])])))).tolist()
    while len(pop) < totalElem:
        pop.append(ran.randint(1,7))
        ran.shuffle(pop)
    pop_list.append(pop)

X=pop_list
pop = Population.new("X", X)
Evaluator().eval(MyProblem_Customized(objs_list=objective_list,cons_list=[2,6,7,10,11]), pop)

main_F=[]
main_X=[]
main_CV=[]
progress_counter = 40;
for i in range(100,200,50):
  x = randfacts.get_fact()
  print (x)
  print ( "We are " + str(progress_counter) + "% done!")
  print (" \n \n \n" )
  progress_counter += 40;
  algorithm = NSGA2(
                  pop_size=i, # converge?
                  n_offsprings=i, # converge?
                  sampling=pop,#get_sampling("int_random"),
                  crossover=SBX(prob=1.0, eta=3.0, vtype=float, repair=RoundingRepair()),
                  mutation=PM(prob=1.0, eta=3.0, vtype=float, repair=RoundingRepair()),
                  eliminate_duplicates=True)
  
  res = minimize(MyProblem_Customized(objs_list=[3,2],cons_list=[2,6,7,10,11]),
               algorithm,
               termination=('n_gen', 100), # n larger Converge 
               seed=1, #  keep seed same
               save_history=True,
               verbose=False,
               return_least_infeasible = True
               )
  #print("Best solution found: %s" % res.X)
  #print("Function value: %s" % res.F)
  #print("Constraint violation: %s" % res.CV)

  if res.CV.all()==0:
    main_F.append(res.F)
    main_X.append(res.X)
    main_CV.append(res.CV)

    check_F=np.concatenate(main_F)
    check_X=np.concatenate(main_X)
    check_CV=np.concatenate(main_CV)

    #print(main_F)
    #print(check_F)
  
  if len(main_F)>0:
    if len(np.unique(check_F,axis=0))>=4:
      break
main_F=np.concatenate(main_F)
main_X=np.concatenate(main_X)
main_CV=np.concatenate(main_CV)




#testCR=X[0][0:totalElem-1]
#testCRTR=X[0][totalElem: (totalElem*2) - 1]
#testR=X[0][totalElem-1]
#testRT=X[0][ (totalElem*2) - 1]

#m_id,cluster_prop,col_prop, uselessElem2 =Layout("linear_adj", y_ax, x_ax, testCR, testCRTR, testR, testRT,matrix_setting,SOA_dict)
#print (m_id)
#view = visualization_matrix(m_id)

#print("Best solution found: %s" % res.X)
#print("Function value: %s" % res.F)
#print("Constraint violation: %s" % res.CV)

F=main_F

X=main_X

def ranking(F):

  fl = F.min(axis=0)
  fu = F.max(axis=0)

  approx_ideal = F.min(axis=0)
  approx_nadir = F.max(axis=0)

  nF = (F - approx_ideal) / (approx_nadir - approx_ideal)

  fl = nF.min(axis=0)
  fu = nF.max(axis=0)

  weights = np.array([0.4, 0.5])

  decomp = ASF()

  d= decomp.do(nF, 1/weights)

  ranking = np.unique(d)
  ranking=ranking.tolist()
  ranking.sort()


  solutions={}
  for i in range(0,len(ranking)):
    index=d.tolist().index(ranking[i])
    solutions[i+1]=F[index].tolist()


  return solutions

def find_five(solution,F):
  selected=[]
  for i in solution:
    select=[]
    searchsol=solution[i]
    sols = np.where(F == searchsol)[0]
    if len(sols)>=10:
      for i in range(0,10,2):
        select.append(sols[i])
    else:
      for i in range(0,len(sols),2):
        select.append(sols[i])
    selected.append(select)

  return selected

solution = ranking(F)
five_sols=find_five(solution,F)

#store value to output as an excel file
iter = len(F)
Fobj=F.tolist()

r_df_Objective1 = {}
r_df_Objective2 = {}
r_df_Column_id = {}
r_df_Cluster_type = {}
r_df_Cluster_ID = {}
r_df_Transform = {}
r_df_Door = {}
r_df_SCWall = {}

for h in range(0,len(five_sols)):
  for i in five_sols[h]:
    #print (i)
    testCR=X[i][0:totalElem]
    testCRTR=np.repeat([0],[totalElem])
    testR=1
    testRT=0
    m_id,cluster_prop,col_prop, uselessElem3=Layout(typology, y_ax, x_ax, testCR.tolist(),testCRTR.tolist(),testR,testRT,matrix_setting,SOA_dict)
    #print(cluster_prop)
    #view = visualization_matrix(m_id)
    r_df_Objective1['Iter_ID'+ str(i)] = Fobj[i][0]
    r_df_Objective2['Iter_ID'+ str(i)] = Fobj[i][1]
    r_df_Column_id['Iter_ID'+ str(i)] = col_prop['Column_ID'] 
    clustername = str()
    valid_transform_list = []

    #find the corresponding cluster_prop to the iteration
    transformation_list = testCRTR.tolist()
    for k in range(len(cluster_prop['Reception'])):
      if cluster_prop['Reception'][k]>0:
        reception_index=k
    #transformation_list.insert(reception_index,testRT)
    
    #Mirror and Unrotate cluster_ID depending on matrix_setting
    ocluster_id = []
    out_SC_wall = []
    if matrix_setting['XY_Mirror'] == True:
      for m in range(len(cluster_prop['Cluster ID'])):
        ocluster_id.append([y_ax - cluster_prop['Cluster ID'][m][0] - cluster_prop['Length'][m] ,cluster_prop['Cluster ID'][m][1]])
      for mw in range(len(SC_wall)):
          SC_wall1 = [y_ax - SC_wall[mw][0][0],SC_wall[mw][0][1]]
          SC_wall2 = [y_ax - SC_wall[mw][1][0],SC_wall[mw][1][1]]
          out_SC_wall.append([SC_wall1,SC_wall2])
    else:
      ocluster_id = cluster_prop['Cluster ID']
      out_SC_wall = SC_wall
    
    if matrix_setting['XY_Rotate'] == True:
      for r in range(len(cluster_prop['Cluster ID'])):
        ocluster_id.append([cluster_prop['Cluster ID'][r][1] + 3,cluster_prop['Cluster ID'][r][0]])
    else:
      ocluster_id = ocluster_id



    #find mid point of the cluster ID
    mcluster_id = []
    if matrix_setting['XY_Rotate'] == True:
      for d in range(len(ocluster_id)):
        midpt_y = ocluster_id[d][0] + (cluster_prop['Breadth'][d]/2)
        midpt_x = ocluster_id[d][1] + (cluster_prop['Length'][d]/2)
        mcluster_id.append([midpt_y,midpt_x])
    else:
      for e in range(len(ocluster_id)):
        midpt_y = ocluster_id[e][0] + (cluster_prop['Length'][e]/2)
        midpt_x = ocluster_id[e][1] + (cluster_prop['Breadth'][e]/2)
        mcluster_id.append([midpt_y,midpt_x])

    clusterid = str()

    for cluster_id in range(len(cluster_prop['Cluster Name'])):
        if cluster_prop['Valid'][cluster_id] == 1:
          if cluster_id < len(cluster_prop['Cluster Name'])-1:
            clustername += (cluster_prop['Cluster Name'][cluster_id]) + ","
            clusterid += str(mcluster_id[cluster_id][0]) + ',' + str(mcluster_id[cluster_id][1]) + '|'
            #valid_transform_list.append(transformation_list[cluster_id])
          else:
            clustername += (cluster_prop['Cluster Name'][cluster_id]) + ","
            clusterid += str(mcluster_id[cluster_id][0]) + ',' + str(mcluster_id[cluster_id][1]) + '|'
            #valid_transform_list.append(transformation_list[cluster_id])
        else:
            pass
    #print (clusterid)
    r_df_Cluster_type['Iter_ID'+ str(i)] = clustername
    

    r_df_Cluster_ID['Iter_ID'+ str(i)] = clusterid


    r_df_Transform['Iter_ID'+ str(i)] = valid_transform_list

    Door_Dir = []
    if matrix_setting['XY_Mirror'] == True:
        Door_Dir.append(0)
    else:
        Door_Dir.append(1)
    
    Door_Dir.append(col_prop['Room Door'][office_col])
    #Append column door location to each cluster in each column
    for room in range(2, len(cluster_prop['Cluster Name']), 1):
        if cluster_prop['Valid'][room] == 1:
            col_num = cluster_prop['Column Index'][room] - 1
            #print (col_num)
            #print (len(col_prop['Room Door']))
            Door_Dir.append(col_prop['Room Door'][col_num])
    r_df_Door['Iter_ID'+ str(i)] = Door_Dir

    SC_wall_pt = str()
    for w in range(len(out_SC_wall)):
      wall_pt1 = str(out_SC_wall[w][0][0]) + ',' + str(out_SC_wall[w][0][1]) + '/'
      wall_pt2 = str(out_SC_wall[w][1][0]) + ',' + str(out_SC_wall[w][1][1]) + '|'
      SC_wall_pt += (wall_pt1 + wall_pt2)
    r_df_SCWall['Iter_ID'+ str(i)] = SC_wall_pt


#output information of whether the matrix is rotated or not
M_rotate = str()
if matrix_setting['XY_Rotate'] == True:
  M_rotate += 'Yes'
else:
  M_rotate += 'No'




revit_dict = {'Objective 1 / sqm':r_df_Objective1, 'Objective 2 / sqm': r_df_Objective2, 'Cluster ID': r_df_Cluster_ID, 'Col ID': r_df_Column_id, 'Cluster Type': r_df_Cluster_type, 'Transform': r_df_Transform, 'Rotate': M_rotate, 'Room Door': r_df_Door, 'Staff Corridor Wall': r_df_SCWall }

r_df = pd.DataFrame(data = revit_dict)

r_df.to_excel('testoutput.xlsx')
writer = pd.ExcelWriter(writepath, engine = 'xlsxwriter')
r_df.to_excel(writer)
writer.save()
writer.close()