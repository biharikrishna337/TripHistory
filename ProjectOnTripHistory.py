import os
import xlrd
import xlsxwriter
import datetime


#Folder Name
outputDir='Result'
courier=outputDir + '/TripResult'
    
#Create folder if not exist
if not os.path.isdir(outputDir):
    os.makedirs(courier)
    print("Home Directory %s was created" %outputDir)

#workbook=xlrd.open_workbook("Tripdata.xlsx")
#worksheet=workbook.sheet_by_name("Sheet1")
#worksheet=workbook.sheet_by_index(0) 
#print("The value at row 4 and column is: {0}".format(worksheet.cell(5,2).value)) 
#print(workbook.row)
#print(workbook.cell_values(0,0))

#file refrenec
inputfile='Tripdata.xlsx'
   

class CapitalBikeShare():
    #Print the all endstaion name but not duplicate and also print their station codes
    def __init__(self):
        global inputfile
        #open inputfile
        self.workbook=xlrd.open_workbook(inputfile)
        self.sheet= self.workbook.sheet_by_index(0)
        #Max active Row
        self.row_count=self.sheet.nrows
        #Max Active Columns
        self.colmn_count=self.sheet.ncols
    
    def wheretheygo(self):
        data=[self.sheet.cell_value(row, 6) for row in range(1,self.sheet.nrows)]
        removeduplicate=set(data)
        return removeduplicate
    

        

#find and store in an excel file the date between start station number and end station number
class WhenDoTheyRide():
    #Mehtod Constructor
    def __init__(self):
        global inputfile
        #open inputfile
        self.workbook=xlrd.open_workbook(inputfile)
        self.sheet= self.workbook.sheet_by_index(0)
        #Max active Row
        self.row_count=self.sheet.nrows
        #Max Active Columns
        self.colmn_count=self.sheet.ncols
#        print(self.colmn_count)
#        print(self.row_count)
            
    def timetheyide(self):
        #sets of staring date and ending date
        #on which time most ride taken
#        for i in range(1,self.row_count):
#            data=[self.sheet.cell_value(i,col) for col in range(1,self.sheet.ncols)]
#            data=[self.sheet.cell_value(i,0)]
        data1=[self.sheet.cell_value(row, 1) for row in range(1,self.sheet.nrows)]
#        print(data1)
        data2=[self.sheet.cell_value(row, 2) for row in range(1,self.sheet.nrows)]
        a=((tuple(zip(data1,data2))))
        return a
        

    
    #check the number of end station on the date

 #function for day on which they reached at the end station for this we need to find the day name of the enddate
 #Find and store of higest duration trip (starting station and ending station)       
class HowFartheygo():
    def __init__(self):
        global inputfile
        #open inputfile
        self.workbook=xlrd.open_workbook(inputfile)
        self.sheet= self.workbook.sheet_by_index(0)
        #Max active Row
        self.row_count=self.sheet.nrows
        #Max Active Columns
        self.colmn_count=self.sheet.ncols
        
        
    def highestduration(self):
#        data=[self.sheet.cell_value(0,col) for col in range(self.sheet.ncols)]
#        data=[self.sheet.cell_value(row1,0) for row1 in range(self.row_count)]
#        print((data))
        data=[self.sheet.cell_value(row1,0) for row1 in range(1,self.row_count)]
        a=("Maximum Duration Spent By The Rider is :",int(max(data)))
#        a= max([cell[0] for cell in self.sheet._cell if cell[1] == column])
        wRow=[self.sheet.cell_value((data.index(max(data))+1),col) for col in range(0,self.colmn_count)]
        heading=[self.sheet.cell_value(0,col) for col in range(0,self.colmn_count)]
#        print(data.index(max(data)))
#        print(wRow)
#        print(heading)
        return (heading),(wRow),(a)



#find the start and end station on the max duration spent
    
    
#finding the most repeated ending station number and display end station name
class PopularStation():
    def __init__(self):
        global inputfile
        #open inputfile
        self.workbook=xlrd.open_workbook(inputfile)
        self.sheet= self.workbook.sheet_by_index(0)
        #Max active Row
        self.row_count=self.sheet.nrows
        #Max Active Columns
        self.colmn_count=self.sheet.ncols
    
    def popular(self):
        data=[self.sheet.cell_value(row1,6) for row1 in range(1, self.row_count)]
        maxtraffic=max(data)
        a=('Most Popular Station: ', maxtraffic,'People Reached at the station', data.count(maxtraffic), 'Times' )
        return (a)
        #check by end station number and print station name
        
    def mostengagedbike(self):
        data=[self.sheet.cell_value(row,7) for row in range(1,self.sheet.nrows)]
#        removeduplicate=(set(data))
        maximum=max(data)
#        print(removeduplicate)
        b=("Most Engage Bike Number is :",maximum)
        return (b)
#

class MostRideOfDay():
    def __init__(self):
        global inputfile
        #open inputfile
        self.workbook=xlrd.open_workbook(inputfile)
        self.sheet= self.workbook.sheet_by_index(0)
        #Max active Row
        self.row_count=self.sheet.nrows
        #Max Active Columns
        self.colmn_count=self.sheet.ncols
    
    def mostride(self):
        m=[]
        data=[self.sheet.cell_value(row1,1) for row1 in range(1,self.row_count)]
#        c=0
        for i in range(len(data)):
            finddate=(data[i].split(' ')[0])
#            c+=1
#            print(c,finddate)  
            month,day,year=(int(x) for x in finddate.split('/'))
            ans=datetime.date(year,month,day)
            dayname=(ans.strftime("%A"))
            m.append(dayname)
            b={}
            for item in m:
                b[item]=b.get(item,0)+1
        return(b)
#        c=max(b)
        
        
#        print("Most ride takend on the Day: ", max(b))
#        for key,val in b.items():
#            print(' \n Most ride taken on the day is: ',key, ' \n and the number is of that day is ',val)
#            print(dayname)
#            print(m)
        
        def peaktimeoftheday():
            pass
            


class ResultData():
    def __init__(self):
        global courier
        self.workbook=xlsxwriter.Workbook(courier+'/Result1.xlsx')
#        self.sheet1= self.workbook.add_worksheet()
##        self.sheet2=self.workbook.add_worksheet()
#        self.sheet3=self.workbook.add_worksheet()
#        self.sheet4=self.workbook.add_worksheet()
#        self.sheet5=self.workbook.add_worksheet()
    def result1(self):
        
#        self.workbook=xlsxwriter.Workbook(courier+'/Result1.xlsx')
        self.sheet1= self.workbook.add_worksheet()
        capital=CapitalBikeShare()
        a=(capital.wheretheygo())
        b=list(a)
#        c=[]
#        print(b)
#        for letter in b:
#            c.append(letter)
        count=1
        for record in b:
            self.sheet1.write(count,0,record)
            count+=1
            self.sheet1.write(0,0,"Where do Capital Bikeshare riders go")
        return record
        self.workbook.close()
        #self.sheet.set_column('A:A',a)
#        worksheet.set_column('A:A',20)
#        bold=workbook.add_format({'bold':True})
#        wb=load_workbook(Workbook(courier+'/Result.xlsx'))
#        ws= wb.active
#        for cell in ws.columns[1]:
#            if cell.value=='abc':
#                print(ws.cell(columnn=12)).value
        
    def result2(self):
        
        self.workbook=xlsxwriter.Workbook(courier+'/Result2.xlsx')
        self.sheet1= self.workbook.add_worksheet()
        wride=WhenDoTheyRide()
        mylist=wride.timetheyide()
        mylist = [str(t) for t in mylist]
        mylist = [str(t) if type(t) is tuple else t for t in mylist]
        self.sheet1.write_column('A1', mylist)


#        print((a))
#        self.sheet1.write_row('A1',[(1,2)])
#        self.workbook=xlrd.open_workbook(courier+'/Result.xlsx/sheet2')
#        self.sheet2= self.workbook.sheet_by_index(1)
#        self.sheet2.write(0,0,'When Do They Ride')
        self.workbook.close()
      
    
    def result3(self):
        hfargo=HowFartheygo()
        heading,wRow,a=hfargo.highestduration()
        self.workbook=xlsxwriter.Workbook(courier+'/Result3.xlsx')
        self.sheet1= self.workbook.add_worksheet()
#        print(a)
        
        a = [str(t) for t in a]
        a = [str(t) if type(t) is tuple else t for t in a]
        self.sheet1.write_row('A1', heading)
        self.sheet1.write_row('A2', wRow)
        self.sheet1.write_row('A3', a)
        print(a)
        self.workbook.close()
#        self.sheet1.write(0,0,a)
        
#        self.workbook = xlrd.open_workbook(courier+'/Result.xlsx')
#        self.sheet3=self.workbook.sheet_by_index(2)
#        self.sheet3.write(1,1,"How far do they go")
#        self.sheet3.write_row(0,0,b)
    
    def result4(self):
        pstation=PopularStation()
        ps1=pstation.popular()
        ps2=pstation.mostengagedbike()
        self.workbook=xlsxwriter.Workbook(courier+'/Result4.xlsx')
        self.sheet1= self.workbook.add_worksheet()
        
#        print(ps1)
#        print(ps2)
        
        ps1 = [str(t) for t in ps1]
        ps1 = [str(t) if type(t) is tuple else t for t in ps1]
        rr1=1
        cc1=0
        for cont in ps1:
#            self.sheet1.write_row('A'+str(rr1),'B'+str(cc1),cont)
            self.sheet1.write(rr1,cc1,cont)
#            rr1+=1
            cc1+=1
        
        ps2 = [str(t) for t in ps2]
        ps2 = [str(t) if type(t) is tuple else t for t in ps2]
        cc1=0
        for cont in ps2:
            rr1+=1
            self.sheet1.write(rr1,cc1,cont)
#            self.sheet1.write_row('A'+str(rr1),'B'+str(cc1),cont)
#            rr1+=1
            cc1+=1
        
        
        self.workbook.close()
#        self.workbook=xlrd.open_workbook(courier+'/Result.xlsx')
#        self.sheet4=self.workbook.sheet_by_index(3)
    
    def result5(self):
        self.workbook=xlsxwriter.Workbook(courier+'/Result5.xlsx')
        self.sheet1= self.workbook.add_worksheet()
        
        mride=MostRideOfDay()
        a=mride.mostride()
        print("Hello ",a)
        cc=1
        rr=1
        for k,v in a.items():
            print(k,v)
            self.sheet1.write('A'+str(rr),k)
            self.sheet1.write('B'+str(cc),v)
            rr+=1
            cc+=1
            
        self.workbook.close()



#capital=CapitalBikeShare()
#print(capital.wheretheygo())
#
#wride=WhenDoTheyRide()
#wride.timetheyide()
#
#find the highest duration betwen start station and end staiton        
#hfargo=HowFartheygo()
#hfargo.highestduration()

#pstation=PopularStation()
#pstation.popular()
#pstation.mostengagedbike()

mride=MostRideOfDay()
mride.mostride()


Result=ResultData()
Result.result1()
Result.result2()
Result.result3()
Result.result4()
Result.result5()
