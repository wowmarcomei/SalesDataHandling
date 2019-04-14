# 列表式定义样例
test = ['A{}'.format(i) for i in range(2,1000+1)]
# print(test)

# 循环定义样例
for i in range(1,1000+1):
    print('A{}'.format(i))

# 字典定义样例
dict1 = {'mei':1,'xu':2,'hong':2}
dict2 = {'mei':1,'xuhong':2,'meixuh':2}

# 字典包含测试样例
if dict2.__contains__('mei'):
    print('yes, I have mei.')

# 迭代字典测试dict1是否存在与dict2中
for key,value in dict1.items():
    if dict2.__contains__(key):
        print(key)
    else:
        print('{} is not included in dict2'.format(key))

# if判断语句赋值
name = "qiwsir" if "laoqi" else "github"        

# 如a = "2013-10-10 23:40:00",想改为 a = "2013/10/10 23:40:00"
# 参考：https://www.cnblogs.com/shgq/p/4065703.html
# b = re.findall(r'\d{4}\/\d{2}\/\d{2}',time.strftime("%Y/%m/%d %H:%M:%S", time.strptime(r_orders_file_sheet['AE9'].value, "%Y-%m-%d %H:%M:%S")))[0]        