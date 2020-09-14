import pandas as pd
import os.path
from pathlib import Path
import datetime

class alter:

    '''
    self.node_dict，key是节点名称，value是拆分后节点名
    self.node_parents key是节点（拆散的），方便用作字符串匹配；value是包含它的节点
    '''
    def __init__(self):
        pass
    def __check_dir(self,path):
        '''
        判断文件路径是否存在
        '''
        myfile = Path(path)

        dir = myfile.is_dir()
        if dir == True:
            print('错误！您输入的是一个文件夹，请输入xlsx的文件路径。')
            return False
        flag = os.path.exists(path)
        if flag == False:
            print(path + ' 文件路径不存在')
            return False
        else:
            return True

    def nan_count(self,List):
        '''
        返回nan有几行，方便用切片切掉
        '''
        # 删除掉nan
        nan_count = 0
        for i in range(len(List)):
            # if List[i] is not True:
            #     nan_count += 1
            if type(List[i]) is not float:
                break
            else:
                nan_count += 1
        # node = node[nan_count::]
        return nan_count


    def check_dir(self):
        if self.__check_dir(self.__node_file) and self.__check_dir(self.__company_file):
            return True
        else:
            return False

    def check_attribute(self):
        '''
        进行节点名称属性的检查
        '''
        self.df1=pd.read_excel(self.__node_file,header=0)
        head1 = list(self.df1.columns)
        flag1 = 0
        for i in head1:
            if '节点名称' == i.strip(' '):
                flag1 = 1
                break
        if flag1 == 0:
            print(self.__node_file+' 文件没有节点名称属性列')
        

        # 企业名称 和 经营范围 检查
        self.df2=pd.read_excel(self.__company_file,header=0)
        head2 = list(self.df2.columns)
        flag2 = 0
        flag3 = 0
        for i in head2:
            if '企业名称' == i.strip(' '):
                flag2 = 1
            elif '经营范围' == i.strip(''):
                flag3 = 1
        if flag2 == 0:
             print(self.__company_file+' 文件没有企业名称属性列')
        
        if flag3 == 0:
             print(self.__company_file+' 文件没有经营范围属性列')

        if flag1*flag2*flag3 == 0:
            return False
        elif flag1*flag2*flag3 == 1:
            return True

    # 节点的包含性检查

    def analyse(self):

        self.node_num = self.df1.shape[0]
        self.company_num = self.df2.shape[0]
        node = list(self.df1['节点名称'])

        node_nan = self.nan_count(node)
        node = node[node_nan::]
        self.node_num -= node_nan # 减去nan的行数

        print('节点信息表的有效数据从%d行开始'%(2+node_nan))
        self.node_offset = 2+node_nan # 行的偏移，为了后期写入数据用


        company = list(self.df2['企业名称'])
        info = list(self.df2['经营范围'])
        company_nan = self.nan_count(company)
        company = company[company_nan::]
        info_nan = self.nan_count(info)
        info = info[info_nan::]

        # 把企业名称和公司名放在一个元组里面
        self.info = []
        for i in range(len(info)):
            self.info.append((company[i],info[i]))


        self.node = node  # 保存一下node，为了进行节点的包含性检查
        # self.info = info #  用不上，就不添加进类的属性
        
        if company_nan == info_nan:
            self.company_num -= company_nan
            print('企业信息表的有效数据从%d行开始'%(2+company_nan))
            self.company_offset = 2+company_nan  # 行的偏移，为了后期写入数据用
        else:
            print('企业名称 和 经营范围的数值的第一行有效值不是同一行，建议退出程序')
        

    def level(self):

        print('节点关联性强度的划分')
        # while(1):
        #     print('希望等分请输入y:')
        #     print('希望自己定义范围输入i:')
        # user_will = input()
        # ack = 'yYiI'

        

        while(1):
            div = input('您想划分成几份>')
            if div.isdigit() == False:
                print('您输入的不是一个整数')
                continue
            if int(div) <= 100:
                break
            print('您输入的已经超过了100')
        div = int(div)
        print('请依次输入'+str(div)+'个关联性的名称，按照关联性从强到弱的顺序输入，以空格分隔。没有输入完成，不允许回车;')
        while(1):
            name = input('请输入节点名称>')
            L = name.split()
            if len(L) == div:
                break
            elif len(L) > div:
                print('您输入的节点名称多了')
            else:
                print('您输入的节点名称少了')
            
        self.conect = L

        self.div_order = []
        M = 100
        ave = int(100/div)
        for i in range(div):
            temp = M
            m = M - ave + 1
            if i != div - 1:
                self.div_order.append((M,m))

            M = m - 1
        self.div_order.append((temp,0))

        for i in range(div):

            print(self.div_order[i],end = ':')
            print(self.conect[i])


        

            


    def analyse_up(self):
        '''
        此函数是本程序的核心
        对于有交叉的节点，我们不做任何处理，直接跳过，我们给出提示让用户去处理。
        '''
        # 还要self.loc获取当前目录的路径
        # warn_txt_loc = self.loc + 'warn.txt'
        warn_txt_loc = self.get_catalogue(self.__node_file)+'warn.txt'
        Now_time = str(datetime.datetime.now()) + ' 更新'
        with open(warn_txt_loc,'w',encoding='utf-8') as f:
            f.write(Now_time)
            f.write('\n')
        company_node = []

        warn_content = []
        # warn_list_split = self.warn_list
        warn_list_split = []
        Letter = r'()（）/、' # 后续有增加在此处增加即可
        for wl in self.warn_list:
            wl0_copy = wl[0]
            wl1_copy = wl[1]

            for w0 in wl[0]:
                if w0 in Letter:
                    wl0_copy = wl0_copy.replace(w0,' ')
            
            for w1 in wl[1]:
                if w1 in Letter:
                    wl1_copy = wl1_copy.replace(w1,' ')

            L0 = wl0_copy.strip().split()
            L1 = wl1_copy.strip().split()
            warn_list_split.append((L0,L1))
            

        # print(warn_list_split)
        # return 


        is_warn = 0
        for i in range(self.company_num):
            company_node.append([])

        node_company_dict = {} # 使用字典，而不是像上一个使用列表是因为有的节点会跳跃，无法确定顺序
        index_i = -1
        # 遍历每个公司的经营范围
        for info_per in self.info:
            info = info_per[1]
            index_i += 1
            index_n = -1

            # 掏出我们已经被打散的节点字典self.node_dict
            # 它的key还是原来的，但是value已经被拆散了，如果value内的单个节点在经营范围内，则把key写入这个企业的节点里面
            for key,value in self.node_dict.items():
                for v in value:
                    # 最短的节点都不在，直接跳过
                    if v not in info:
                        continue

                    else:
                        
                        # 如果节点在交叉域内则跳过，不做任何处理
                        warn_flag = 1
                        content = ''
                        # for w in self.warn_list: # 建议把warn_list拆散
                        for w in warn_list_split:
                            if v in w[0] and v in w[1]:
                                is_warn = 1

                                # self.warn_user_decide 已经转化成int型，不用担心
                                warn_flag = int(self.warn_user_decide) # 0,跳过，1 全部写入

                                # 决定交叉的是否写入，如果不要写入的话，注释掉这一行
                                con = str(info_per[0])+':'+v+'->'+str(w[0])+'和'+str(w[1])
                                con2 = str(info_per[0])+':'+v+'->'+ str(w[1]) + '和' + str(w[0])
                                if con not in warn_content and con2 not in warn_content:
                                    warn_content.append(con)                                                            

                        
                        
                        
                        if warn_flag == 0:
                            continue
                        
                        info_copy_for_index = info.strip() # 万一前后有多余的空格进行去除

                        # 没有包含它的节点
                        if len(self.node_parents[v]) == 0:

                            # 找到这个节点在经营范围的结束位置
                            index_node = info_copy_for_index.index(v) # 这个拆分后的节点而不是原始的节点

                            # 企业信息的填充部分
                            if len(info_copy_for_index)-len(v) == 0:
                                occupy = 100
                                occupy = str(occupy)
                                node_level = self.conect[0]

                            else:
                                # +0.5是为了进行4舍5入，保留整数部分
                                # 此处的index_node不需要加一，下标从0开始
                                occupy = index_node/(len(info_copy_for_index)-len(v))*100
                                occupy = 100 - occupy
                                occupy += 0.5
                                occupy = str(occupy).split('.')
                                occupy = occupy[0]
                                div_order_count = 0
                                for i in self.div_order:
                                    if int(occupy)<= i[0] and int(occupy) >= i[1]:
                                        node_level = self.conect[div_order_count]
                                        break
                                    div_order_count += 1
                                    
                                


                            add_str = '('+str(index_node+1) + '-' + str(len(info_copy_for_index))+'-'+occupy+'%-'+ node_level +')'


                            # 公司添加这个节点
                            if key+add_str not in company_node[index_i]:
                                company_node[index_i].append(key+add_str) # 注意此处是append，原始的节点
                            # 节点添加这个公司,字典是否有这个键
                            if key not in node_company_dict:
                                node_company_dict[key] = []
                                node_company_dict[key].append(info_per[0])
                            else:
                                # 字典是否已经有了这个值
                                if info_per[0] not in node_company_dict[key]:
                                    node_company_dict[key].append(info_per[0])


                        else:
                            # 就一个个找，找不到再添加原始的节点
                            # 先备份一份info
                            flag = 0
                            info_copy = info
                            for p_value in self.node_parents[v]: # p_value是一个元组(拆分的节点，它的原始节点名)

                                # 包含短的节点的长的节点的查找
                                if p_value[0] in info_copy:
                                    # 循环切除
                                    while(1):
                                        info_copy = info_copy.replace(p_value[0],' ') # 节点寻找开始了
                                        if p_value[0] not in info_copy:
                                            break

                                    
                                    # 找到这个节点在经营范围的结束位置
                                    index_node = info_copy_for_index.index(p_value[0]) # 这个拆分后的节点而不是原始的节点
                                    
                                    # 企业信息的填充部分
                                    if len(info_copy_for_index)-len(p_value[0]) == 0:
                                        occupy = 100
                                        occupy = str(occupy)
                                        node_level = self.conect[0]

                                        
                                    else:
                                        # 
                                        occupy = index_node/(len(info_copy_for_index)-len(p_value[0]))*100
                                        occupy = 100 - occupy
                                        occupy += 0.5
                                        occupy = str(occupy).split('.')
                                        occupy = occupy[0]
                                        div_order_count = 0
                                        for i in self.div_order:
                                            if int(occupy)<= i[0] and int(occupy) >= i[1]:
                                                node_level = self.conect[div_order_count]
                                                break
                                            div_order_count += 1


                                    # index_node的下标，从0开始，需要对他加1
                                    add_str = '('+str(index_node+1) + '-' + str(len(info_copy_for_index))+'-'+occupy+'%-'+node_level+')'


                                    
                                    # 对这个公司添加这个原始节点
                                    # p_value[1]是节点的原始节点名字
                                    # 已更新为对填充部分也进行比对，不会出现上一个版本的重复现象
                                    if p_value[1]+add_str not in company_node[index_i]: 
                                        company_node[index_i].append(p_value[1]+add_str)
                                    # 对这个节点添加这个公司
                                    # info_per[0]是公司名字
                                    if p_value[1] not in node_company_dict:
                                        node_company_dict[p_value[1]] = []
                                    if info_per[0] not in node_company_dict[p_value[1]]:
                                        node_company_dict[p_value[1]].append(info_per[0])

                            # 所有比它长度要长的节点，都切掉了，那么现在我们可以用它本身来进行匹配，此时不存在任何的干扰
                            # 也是写入文件的地方，特别容易被遗漏，

                            if v in info_copy:
                                # 找到这个节点在经营范围的结束位置
                                index_node = info_copy_for_index.index(v) # 这个拆分后的节点而不是原始的节点
                                    
                                # 企业信息的填充部分
                                if len(info_copy_for_index)-len(v) == 0:
                                    occupy = 100
                                    occupy = str(occupy)
                                    node_level = self.conect[0]

                                        
                                else:
                                        # 
                                    occupy = index_node/(len(info_copy_for_index)-len(v))*100
                                    occupy = 100 - occupy
                                    occupy += 0.5
                                    occupy = str(occupy).split('.')
                                    occupy = occupy[0]
                                    div_order_count = 0
                                    for i in self.div_order:
                                        if int(occupy)<= i[0] and int(occupy) >= i[1]:
                                            node_level = self.conect[div_order_count]
                                            break
                                        div_order_count += 1
                                
                                # index_node的下标，从0开始，需要对他加1
                                add_str = '('+str(index_node+1) + '-' + str(len(info_copy_for_index))+'-'+occupy+'%-'+node_level+')'


                                # v的原始节点名字叫key
                                if v+add_str not in company_node[index_i]:
                                    company_node[index_i].append(key+add_str)
                                # v属于的公司名字叫什么
                                if key not in node_company_dict:
                                    node_company_dict[key] = []
                                if info_per[0] not in node_company_dict[key]:
                                        node_company_dict[key].append(info_per[0])


        if is_warn == 1:
            with open(warn_txt_loc,'a+',encoding='utf-8') as f:
                f.write('\n'.join(warn_content))
                f.write('\n')
            
            print()
            print('【注意】:已经为您生成warn.txt文件，里面的节点是存在交叉的节点。')
            print('warn.txt路径为:'+warn_txt_loc)
            if self.warn_user_decide == '1':
                print('您选择的是全部写入模式，为您把存在交叉的节点进行全部写入。')
            elif self.warn_user_decide == '0':
                print('您选择的是跳过模式，对于warn.txt里面存在交叉的节点，程序不会对它们进行任何处理。')
                print('您可以考虑进行修改节点，使之没有交叉，或者考虑进行人工操作来决定写入哪个节点。')
                
        # print(len(node_company_dict))
        self.node_company = []
        for nd in self.node:
            # if len(node_company_dict[nd]) == 0:
            #     self.node_company.append([])
            if nd not in node_company_dict:
                 self.node_company.append('')
            else:
                self.node_company.append('、'.join(node_company_dict[nd]))

        # print(company_node)
        # company_node没问题

        self.company_node = []
        for c_n in company_node:
            if len(c_n) == 0:
                self.company_node.append('')
            else:
                self.company_node.append('&'.join(c_n))

    def node_split(self):
        # 允许划分节点的符号集合
        Letter = r'()（）/、' # 后续有增加在此处增加即可

        self.node_dict = {}
        for i in self.node:
            if i in self.node_dict:
                print('警告！出现错误！'+i+'节点名称重复')
            else:
                # 拆分开始
                s = i.strip()

                # 节点间隔符号replace成空格
                
                for j in s:
                    if j in Letter:
                        s = s.replace(j,' ')

                self.node_dict[i] = s.split()

    def check_node(self):
        '''
        只检查短的节点是否和长的节点拆分后的节点重复
        因为节点长度长的节点是OK的，长度短的会被长度长的节点所包含
        把所有的节点拆分之后，混在一个字典中
        生成属性self.node_parents # 字典
        对它的value，已经按照字符串的长度进行了升序排列（因为长的绝对是OK的，短的可能就不OK了，所以把长的放在前面）
        '''
 
        warn_list = []
        warn_flag = 0
        self.node_parents = {}
        
        # for i in self.node:
        count1= -1
        for key,value in self.node_dict.items():
            count1 += 1

            # 对第一个的节点进行拆分
            for per in value:
                count2 = -1

                # 拆散节点，作为键
                if per not in self.node_parents:
                    self.node_parents[per] = []

                for j_key,j_value in self.node_dict.items():
                    count2 += 1
                    # 同一个节点，就是它本身直接跳过
                    if count1 == count2:
                        continue
                    
                    # 查节点

                    for k in j_value:

                        # 交叉性检查
                        if per == k:
                            warn_flag = 1
                            warn_list.append((key,j_key))

                        # 包含性检查,显然也不能和它自己相同
                        elif per in k:
                            # 需要添加元组进去，因为最后我们是添加它的原始节点，而不是拆分后的节点
                            if (k,j_key) not in self.node_parents[per]:
                                self.node_parents[per].append((k,j_key))
                            # 就是看k，在不在里面避免重复添加
                            # (k,parent)在不在里面
                            # if k not in self.node_parents[per]:
                            #     
                            #     self.node_parents[per].append(k)

        for key,value in self.node_parents.items():
            self.node_parents[key] = sorted(value,key = lambda i:len(i[0]),reverse=True)
        
        self.warn_list = warn_list

        if warn_flag == 1:
            print('警告！节点和节点拆分后的节点重复，出现交叉！在下面已经列出')
            # print(warn_list)
            for wl in warn_list:
                print(wl)
             
            print('对于上述这些节点，程序无法决定添加哪一个节点。')
            while(1):
                print('若希望程序遇到交叉节点，进行跳过，请输入 0:')
                print('若希望程序遇到交叉节点，把交叉的节点全部写进去，请输入 1:')
                self.warn_user_decide = input('>')
                if self.warn_user_decide == '0' or self.warn_user_decide == '1':
                    break

    def get_catalogue(self,loc):
        for i in range(len(loc)):
            if loc[-i] == '\\':
                return loc[:-i+1:]
        return ''


    def out_filename(self,loc):
        '''
        C:/xx/某某.xlsx -> C:/xx/某某new.xlsx
        '''        
        for i in range(1,len(loc)):
            # 找到.
            if loc[-i] == '.':
                loc = loc[:-i:] + 'new' + loc[-i::]
                self.loc = loc
                return loc

    def write_excel(self):
        outfile1 = self.out_filename(self.__node_file)
        self.df1['企业名称'] = ['']*(self.node_offset-2)+self.node_company
        # node_index = [i+self.node_offset for i in range(len(self.node_company))]
        # print(node_index)
        self.df1.to_excel(outfile1,index = None)
        print(outfile1,' 输出完成')

        outfile2 = self.out_filename(self.__company_file)
        self.df2['节点名称'] = ['']*(self.company_offset-2)+self.company_node

        # company_index = [i+self.company_offset for i in range(len(self.company_node))]
        # print(company_node)
        # print(self.company_node)
        self.df2.to_excel(outfile2,index=None)
        print(outfile2,' 输出完成')

    def Main(self):
        # 得用input
        print('如果exe文件和excel表在同一个目录下只需要输入文件名即可')
        print('不在同一个目录下，必须输入路径+文件名(绝对路径)')
        print('程序只支持xlsx文件，不支持xls文件')
        self.__node_file = input('输入节点信息表的绝对路径>')
        self.__company_file = input('输入企业信息表的绝对路径>')
        self.live = self.check_dir()

        if self.live:
            self.live = self.check_attribute()
        else:
            print('bye')
            return
        if self.live:
            print('数据分析中...')
            self.level()
            self.analyse()
            self.node_split()
            self.check_node()
            self.analyse_up()
            print('数据分析完成')
            print('正在写入新的文件...')
            self.write_excel()
            print('本次操作完成！\n输出的文件和源文件在同一个文件夹下面')
            start = input('输入y,继续运行下一个文件;输入q，退出程序。\n>')
            if start == 'y' or start == 'Y':
                self.Main()
            elif start == 'q' or start == 'Q':
                return 


if __name__ == '__main__':

    # excel = alter('02-节点信息表.xlsx','03-企业信息表.xlsx')
    excel = alter()
    excel.Main()
    # excel.level()


