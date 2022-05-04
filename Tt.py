import win32com.client as win32#library


CANoe=win32.DispatchEx("CANoe.Application")#Loads canoe

CANoe.Open("C:\\Users\\Public\\Documents\\Vector\\CANoe\\10.0\\CANoe Sample Configurations\\CAN\\TestFeatureSet\\CentralLockingSystem\\CentralLockingSystem.cfg")#loads configuration

CANoe.Measurement.Start()#run the configuration
testSetup=CANoe.Configuration.TestSetup#Access the testSetup
 
testEnv=testSetup.TestEnvironments.Item(1)#Accessing the testEnvironment
testModules=testEnv.TestModules#Accessing Test modules
for i in range(1,testModules.Count+1):#read all test modules
    test_Module=testEnv.TestModules.Item(i)#Accessing the first testModule
    print(test_Module.Name)#printing the name of testModule
     
    seq=test_Module.Sequence#Accessing test Sequence
     
    for i in range(1,seq.Count+1):#Accessing test case
        testCase=win32.CastTo(seq.Item(i),"")#type casting test case
        #print(testCase.Name)#printing testcase
     
    test_Module.Start()#starting testcase execution
