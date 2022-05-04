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

    
    
    
    import win32com.client as win32

if __name__ == '__main__':
    
    '''Launch'''
    CANoe = win32.DispatchEx("CANoe.Application")
    '''Load Configuration'''
    CANoe.Open("C:\\Users\\Public\\Documents\\Vector\\CANoe\\10.0\\CANoe Sample Configurations\\CAN\\TestFeatureSet\\CentralLockingSystem\\CentralLockingSystem.cfg")
    '''Start Measurement'''    
    CANoe.Measurement.Start()
    
    systemCAN = CANoe.System.Namespaces
    sys_namespace = systemCAN("SystemUnderTest")
    sysVariables = sys_namespace.Variables
    print(type(sysVariables))
    
    for i in range(1, sysVariables.Count+1):
        sysVar = win32.CastTo(sysVariables.Item(i), "IVariable")
        print(sysVar.Name)
   
    sys_value = sysVariables("ErrorInCrashSensorUsage")
    sys_value.Value = 1
    
    sys_value2 = sys_namespace.Variables("ErrorCrashSensorOnVelocity")
    sys_value2.Value = 1
    
    signalValue = CANoe.GetBus("CAN").GetSignal(1, "VehicleMotion", "EngineRunning")
    engineRunning = win32.CastTo(signalValue, "ISignal4")
    print(engineRunning.Value)
    
#     print(type(CANoe.Environment.getVariable("").Value)
    
    while engineRunning.Value == 0:
        print(CANoe.GetBus("CAN").GetSignal(1,"VehicleMotion","EngineRunning"))

    systemCAN = CANoe.System.Namespaces
    sys_namespace = systemCAN("SystemUnderTest")
    sys_value = sysVariables("ErrorInCrashSensorUsage")
    print(sys_value.Value)
    
    testSetup = CANoe.Configuration.TestSetup
    test_env = testSetup.TestEnvironments.Item(1)
    test_env = win32.CastTo(test_env, "ITestEnvironment2")
    print(test_env.Name)
    
    testModules = test_env.TestModules
    result = 0
    for i in range(1, testModules.Count+1):
        test_module = test_env.TestModules.Item(i)
        print(test_module.Name)
        
        seq = test_module.Sequence
        for i in range(1, seq.Count + 1):
            tc = win32.CastTo(seq.Item(i), "ITestCase")
            print(tc.Name)
            
        test_module.Start()
    
    
