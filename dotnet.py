import clr
import sys
import System  # 直接使用 .NET System 命名空间
from System.Threading import Thread, ThreadStart, ApartmentState

# 设置C# DLL文件路径，假设与Python文件在同一文件夹中
sys.path.append(r'pythonProject')

# 加载C# DLL
clr.FindAssembly('PythonCallDll.dll')
clr.AddReference('PythonCallDll')  # 加载引用

# 导入C#命名空间
from PythonCallDll import *

def run_form():
    instance = Class1()  # 创建类的对象实例1，使用无参构造器
    instance2 = Class1('object2')  # 创建类的对象实例2，使用有参构造器
    print('Class1类的静态字段(对类的实例化次数计数)_classObjectCounts = ', Class1._classObjectCounts)
    print('Class1类的实例1，调用方法AddCalc: 1+2= ', instance.AddCalc(1, 2))
    print('Class1类的实例1，实例属性ClassName:', instance.ClassName)
    print('Class1类的实例2，实例属性ClassName:', instance2.ClassName)

    try:
        # 定义变量以传递给SayHello方法
        device_name = "DS-970"
        file_path = "c:\\debug"
        image_type = 0
        doc_source = 2
        resolution =200  #dpi
        skewCorrect =2 # 矫正偏差

        docsize =2 #A4
        fileformat = 1 # 1 jpeg  2
        logPath="c:\\debug"
        # 调用有参SayHello方法，传递动态参数
        form = Form2()  # 获取 dll 下 form类的实例
        form.SayHello(device_name, file_path, image_type, doc_source,resolution,skewCorrect,docsize, fileformat,logPath)
    except Exception as e:
        print(f"An error occurred: {e}")

# 创建一个线程并将其设置为STA
thread = Thread(ThreadStart(run_form))
thread.SetApartmentState(ApartmentState.STA)
thread.Start()
thread.Join()  # 等待线程完成
