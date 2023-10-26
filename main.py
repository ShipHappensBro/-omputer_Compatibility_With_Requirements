import winapps,wmi,psutil,openpyxl,os,socket,GPUtil

class CompatibilityWithAutocad():

    def __init__(self) -> None:
        """
        Атрибут `self.AUTOCADS` - это список словарей, содержащих информацию о требованиях различных версий AutoCAD к аппаратному обеспечению.

        Атрибут `self.SOLID_EDGE` - это словарь, содержащий информацию о требованиях Solid Edge к аппаратному обеспечению.

        Атрибут `self.DIRECTORY` - это путь к файлу Excel, в который будет записана информация о совместимости.

        Атрибуты `self.username` и `self.hostname` получают имя пользователя и имя хоста соответственно.

        Атрибут `self.COMPUTER` представляет собой экземпляр класса WMI, который используется для получения информации о системе.

        Атрибуты `self.computer_info`, `self.cpu_ghz`, `self.ram`, `self.cpu_name`, `self.cpu_info` и `self.gpu_name` получают информацию о системе, такую как информация о операционной системе, частота процессора, объем оперативной памяти, имя процессора и имя графического процессора.

        """
        self.AUTOCADS=[{'autocad lt 2024':{'cpu':3,'ram':8,'videoram':1}},
                       {'autocad lt 2017':{'cpu':1.5,'ram':4,'videoram':1}},
                       {'autocad lt 2012':{'cpu':1.4,'ram':4,'videoram':1}},
                       {'autocad 2024':{'cpu':3,'ram':8,'videoram':2}},
                       {'autocad 2023':{'cpu':2.5,'ram':8,'videoram':1}},
                       {'autocad 2017':{'cpu':2,'ram':4,'videoram':1}},
                       {'autocad 2016':{'cpu':2,'ram':4,'videoram':1}},
                       {'autocad 2015':{'cpu':3,'ram':4,'videoram':1}},
                       {'autocad 2012':{'cpu':3,'ram':4,'videoram':1}}]
        
        self.SOLID_EDGE={'cpu':2.5,'ram':16,'ram2':8,'videoram':1}
        self.DIRECTORY='C:\\inv\\compatibility.xlsx'
        self.username=os.getlogin()
        self.hostname=socket.gethostname()
        self.COMPUTER=wmi.WMI()
        self.computer_info=self.COMPUTER.Win32_OperatingSystem()[0].name.split("|")[0]
        self.cpu_ghz=round(psutil.cpu_freq().max/1024,1)
        self.ram=round(psutil.virtual_memory().total/1024/1024/1024,0)
        self.cpu_name=self.COMPUTER.Win32_Processor()[0].name
        self.cpu_info=round(self.cpu_ghz,1)
        self.gpu_name=self.COMPUTER.Win32_VideoController()[0].name

    def __str__(self) -> str:

        print(f'\n\nWindows: {self.computer_info}\nCPU: {self.cpu_name}\nCPU_GHz: {self.cpu_info}@GHz\nGPU: {self.gpu_name}\nGPU_RAM: {self.gpu_ram} Gb\nPhysical ram: {self.ram} Gb')

    def gpu(self) -> None:
        """
        Функция `gpu` используется для получения информации о графическом процессоре (GPU) на компьютере.Работает только для видеокарт Nvidia
        
        Nvidia: использует `nvidia-smi` ВАЖНО

        Сначала функция пытается получить список доступных GPU с помощью `GPUtil.getGPUs()`. Если GPU не обнаружены или происходит исключение, переменная `gpus` устанавливается в `False`.

        Если `gpus` не равно `False`, функция перебирает каждый GPU в списке и получает имя видеоконтроллера и объем памяти GPU.

        Если `gpus` равно `False`, функция пытается получить имя видеоконтроллера и объем памяти видеоконтроллера с помощью `Win32_VideoController()`. Если это не удается, функция устанавливает объем памяти в 0.

        Аргументы:
            self: экземпляр класса.

        Возвращает:
            Функция не возвращает значения, но устанавливает значения атрибутов `gpu_name` и `gpu_ram`.
        """
        try:
            #Работает только для видеокарт Nvidia
            gpus = GPUtil.getGPUs()
            if gpus == []:
                gpus=False
        except:
            gpus=False

        if gpus!=False:
            for gpu in gpus:
                self.gpu_name=self.COMPUTER.Win32_VideoController()[0].name
                self.gpu_ram = gpu.memoryTotal/1024
        else:
            #Для видеокарт других производителей,или встроенной графики
            try:
                self.gpu_name=self.COMPUTER.Win32_VideoController()[0].name
                self.gpu_ram=self.COMPUTER.Win32_VideoController()[0].AdapterRAM/1024/1024/1024
            except:
                self.gpu_name=self.COMPUTER.Win32_VideoController()[0].name
                self.gpu_ram=self.COMPUTER.Win32_VideoController()[0].AdapterRAM
                if self.gpu_ram==None:
                    self.gpu_ram=0

    def scan_autocad(self) -> None:
        """
        Функция `scan_autocad` используется для сканирования установленных приложений AutoCAD на компьютере и записи информации о них в файл Excel.

        Сначала функция проверяет, установлено ли приложение AutoCAD на компьютере. Если AutoCAD не обнаружен, функция возвращает `False`.

        Если AutoCAD обнаружен, функция пытается открыть рабочую книгу Excel по указанному пути. Если рабочая книга не может быть открыта, функция выводит сообщение об ошибке и возвращает `False`.

        Затем функция перебирает каждое приложение в списке `AUTOCADS` и записывает информацию о нем в новую строку в листе 'Autocad' рабочей книги. Информация включает имя хоста, информацию о компьютере, имя AutoCAD, имя процессора, объем оперативной памяти, имя графического процессора и объем видеопамяти.

        Наконец, функция проверяет, соответствует ли текущая конфигурация системы требованиям AutoCAD. Если система соответствует требованиям, она записывает "Совместимо" в ячейку и окрашивает ее в зеленый цвет. В противном случае она записывает "Несовместимо" и окрашивает ячейку в красный цвет.

        Аргументы:
            self: экземпляр класса.

        Возвращает:
            Функция не возвращает значения, но сохраняет изменения в рабочей книге Excel.
        """
        autocad=False
        for app in winapps.list_installed():
            if 'autocad' in app.name.lower():
                autocad=app.name.lower()
        if autocad==False:
            return(False)
        try:
            wb = openpyxl.load_workbook(self.DIRECTORY)
            ws = wb['Autocad']
        except:
            print(f'Не удалось открыть {self.DIRECTORY}')
            return(False)
        last_row=ws.max_row+1
        for app in self.AUTOCADS:
            key=list(app.keys())[0]
            if not key in autocad:
                continue
            dict_req=list(app.values())[0]
            cpu_req=dict_req['cpu']
            ram_req=dict_req['ram']
            videoram_req=dict_req['videoram']

            ws.cell(row=last_row, column=1).value = self.hostname
            ws.cell(row=last_row, column=2).value = self.hostname
            ws.cell(row=last_row, column=3).value = self.computer_info
            ws.cell(row=last_row, column=4).value = autocad
            ws.cell(row=last_row, column=5).value = self.cpu_name
            ws.cell(row=last_row, column=6).value = self.ram
            ws.cell(row=last_row, column=7).value = self.gpu_name
            ws.cell(row=last_row, column=8).value = self.gpu_ram
            if self.cpu_ghz>=cpu_req and self.ram>=ram_req and self.gpu_ram>=videoram_req:
                ws.cell(row=last_row, column=9).value = "Совместимо"
                ws.cell(row=last_row, column=9).fill = openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
            else:
                ws.cell(row=last_row, column=9).value = "Несовместимо"
                ws.cell(row=last_row, column=9).fill = openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        wb.save(self.DIRECTORY)

    def scan_solid(self) -> None:
        """
        Функция `scan_solid` используется для сканирования установленных приложений Solid Edge на компьютере и записи информации о них в файл Excel.

        Сначала функция проверяет, установлено ли приложение Solid Edge на компьютере. Если Solid Edge не обнаружен, функция возвращает `False`.

        Если Solid Edge обнаружен, функция пытается открыть рабочую книгу Excel по указанному пути. Если рабочая книга не может быть открыта, функция выводит сообщение об ошибке и возвращает `False`.

        Затем функция записывает информацию о системе и приложении Solid Edge в новую строку в листе 'Solid' рабочей книги. Информация включает имя хоста, информацию о компьютере, имя Solid Edge, имя процессора, объем оперативной памяти, имя графического процессора и объем видеопамяти.

        Наконец, функция проверяет, соответствует ли текущая конфигурация системы требованиям Solid Edge для полной и учебной версий. Если система соответствует требованиям, она записывает "Совместимо с полной версией" или "Совместимо с учебной версией" в соответствующую ячейку и окрашивает ее в зеленый цвет. В противном случае она записывает "Несовместимо с полной версией" или "Несовместимо с учебной версией" и окрашивает ячейку в красный цвет.

        Аргументы:
            self: экземпляр класса.

        Возвращает:
            Функция не возвращает значения, но сохраняет изменения в рабочей книге Excel.
        """
        solid_edge=False
        for app in winapps.list_installed():
            if 'solid' in app.name.lower():
                solid_edge=app.name.lower()
        if solid_edge==False:
            return(False)    
        try:
            wb = openpyxl.load_workbook(self.DIRECTORY)
            ws = wb['Solid']
        except:
            print(f'Не удалось открыть {self.DIRECTORY}')
            return(False)
        last_row=ws.max_row+1
        cpu_req=self.SOLID_EDGE['cpu']
        ram_for_full=self.SOLID_EDGE['ram']
        ram_for_lesson=self.SOLID_EDGE['ram2']
        gpu_req=self.SOLID_EDGE['videoram']
        ws.cell(row=last_row, column=1).value = self.hostname
        ws.cell(row=last_row, column=2).value = self.hostname
        ws.cell(row=last_row, column=3).value = self.computer_info
        ws.cell(row=last_row, column=4).value = solid_edge
        ws.cell(row=last_row, column=5).value = self.cpu_name
        ws.cell(row=last_row, column=6).value = self.ram
        ws.cell(row=last_row, column=7).value = self.gpu_name
        ws.cell(row=last_row, column=8).value = self.gpu_ram
        if self.cpu_ghz>=cpu_req and self.ram>=ram_for_full and self.gpu_ram>=gpu_req:
            ws.cell(row=last_row, column=9).value = "Совместимо с полной версией"
            ws.cell(row=last_row, column=9).fill = openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
        else:
            ws.cell(row=last_row, column=9).value = "Несовместимо с полной версией"
            ws.cell(row=last_row, column=9).fill = openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')

        if self.cpu_ghz>=cpu_req and self.ram>=ram_for_lesson and self.gpu_ram>=gpu_req:
            ws.cell(row=last_row, column=10).value = "Совместимо с учебной версией"
            ws.cell(row=last_row, column=10).fill = openpyxl.styles.PatternFill(start_color='00FF00', end_color='00FF00', fill_type='solid')
        else:
            ws.cell(row=last_row, column=10).value = "Несовместимо с учебной версией"
            ws.cell(row=last_row, column=10).fill = openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
        wb.save(self.DIRECTORY)

    def main(self) -> None:
        """
        Функция `main` является основной функцией класса, которая вызывает другие функции для сбора информации о системе и установленных приложениях.

        Сначала функция вызывает `self.gpu()`, чтобы получить информацию о графическом процессоре (GPU) на компьютере.

        Затем функция вызывает `self.scan_autocad()`, чтобы сканировать установленные приложения AutoCAD на компьютере и записывать информацию о них в файл Excel.

        Наконец, функция вызывает `self.scan_solid()`, чтобы сканировать установленные приложения Solid Edge на компьютере и записывать информацию о них в файл Excel.

        Аргументы:
            self: экземпляр класса.

        Возвращает:
            Функция не возвращает значения.
        """
        self.gpu()
        self.scan_autocad()
        self.scan_solid()

if __name__=="__main__":
    compatibility=CompatibilityWithAutocad()
    compatibility.main()
