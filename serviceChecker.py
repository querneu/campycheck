#Para instalar o servico, utilize o comando:
#     python serviceChecker.py install
#Para iniciar o servico, utilize o comando:
#     NET START campycheck
#Para parar o servico, utilize o comando:
#     NET STOP campycheck

# REQUIRED: https://github-production-release-asset-2e65be.s3.amazonaws.com/108187130/358fd300-06ca-11ea-85a9-501fe287636c?X-Amz-Algorithm=AWS4-HMAC-SHA256&X-Amz-Credential=AKIAIWNJYAX4CSVEH53A%2F20200116%2Fus-east-1%2Fs3%2Faws4_request&X-Amz-Date=20200116T122709Z&X-Amz-Expires=300&X-Amz-Signature=455dd5079df3777f734d4ea47c3002ef1fa4bc566c72fdd26971a86742b030da&X-Amz-SignedHeaders=host&actor_id=47905444&response-content-disposition=attachment%3B%20filename%3Dpywin32-227.win-amd64-py3.7.exe&response-content-type=application%2Foctet-stream

import win32service
import win32serviceutil
import win32event
import servicemanager
import os 
import datetime
import ctypes
import requests
import glob

save_path = "C:/Users/Lucas.Leite/Desktop/campycheck/" #mude o caminho de acordo com o seu ambiente

filename = "RequestLog.csv"
completePath = os.path.join(save_path,filename)
arquivo = open(completePath, "w+")

class ServiceChecker(win32serviceutil.ServiceFramework):
    # Voce pode utilizar NET START/STOP com o nome do servico a seguir
    _svc_name_ = "campycheck"
    # Texto que mostre o nome do servico no scm
    _svc_display_name_ = "Campaign Service Verifier"
    # Descricao do servico
    _svc_description_ = "Servico de verificacao de campanhas "
   
    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self,args)
        # Cria um evento para ouvir as requisicoes de NET STOP
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
    # logica central do servico, aqui voce cria os codigos para executar   
    
    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_,''))
        self.main()

    def SvcStop(self):
        # diz ao scm que esta sendo desligado
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # inicia o evento de parada
        win32event.SetEvent(self.hWaitStop)

    def main(self):

        def renameAllFiles():

            for i, filename in enumerate(glob.glob(save_path+"*.txt")):
                os.rename(filename, os.path.join(save_path, filename+".bak"))

        def getFiles():
            user =  "administrator@labcce.com" #Apague antes de comitar
            password = "brzPWD@16" #Apague antes de comitar#
            campaignId = 5005
            os.chdir(save_path)

            for currentFile in glob.glob("*.txt"):
                logIT("Arquivo atual: "+str(currentFile))

                if '.txt' in currentFile:
                    campaignFile = {'upload_file': open(save_path+currentFile, 'rb')}
                    url = "https://192.168.110.152/unifiedconfig/config/campaign/"+str(campaignId)+"/import"
                    request = requests.get(
                        url,verify=False, 
                        auth=(user,password), 
                        data=campaignFile,
                        headers={
                            "content-type":"text/plain"
                        })

                    if request.status_code == 200:
                        logIT(str(now)+currentFile+"Success!")

                    else:
                        logIT(str(now)+str(request.status_code))
                        logIT(str(now)+str(currentFile)+".bak"+"----Fail")

        def logIT(msg):
            #Escrevendo informações no log
            now = datetime.datetime.now()
            arquivo.write(str(now)+" - "+msg+"\n")
            arquivo.flush()
            #Espera 5 segundos para escrever a nova informacao no log
    
        try:##Tentar realizar procedimentos durante o tempo da execução
            rc = None
            while rc != win32event.WAIT_OBJECT_0:
                logIT(str(now)+"Entrou no metodo principal") #Aqui você coloca a mensagem do log para o que você quiser, operação do log sera feita a cada chamada
                getFiles()
                renameAllFiles()
                rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)
            arquivo.write(str(datetime.datetime.now())+" - Fechando serviço\n")
            arquivo.close()
            #ctypes.windll.user32.MessageBoxW(0, "Log Criado com sucesso em: "+completePath, "Mensagem", 0)
        
        except IndexError: ##Pegar erro na saida do terminal
            traceback.print_exception(exc_type, exc_value, exc_traceback,
                              limit=2, file=sys.stdout)

if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(ServiceChecker)