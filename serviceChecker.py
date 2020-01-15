#Para instalar o servico, utilize o comando:
#     python serviceChecker.py install
#Para iniciar o servico, utilize o comando:
#     NET START campycheck
#Para parar o servico, utilize o comando:
#     NET STOP campycheck
import win32service
import win32serviceutil
import win32event
import servicemanager
import os #aplicacoes do sistema
import datetime

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
        

        #Altere o caminho do usuario, pode ser feito com os.environ.keys() para receber as chaves disponiveis no ambiente
        #pode ser dado output para qualquer tipo de arquivo
        #substitua o filepath ate o diretorio do srvcFileChecker para onde desejares que fiquem salvos os logs
        #delimitado na os.environ["chave"]
        save_path = "C:/Users/Lucas.Leite/Desktop/srvcFileChecker"
        #nome do arquivo de log
        filename = "RequestLog.txt"
        #caminho completo
        completePath = os.path.join(save_path,filename)
        #Se o evento de stop (NET STOP SERVICE) nao for gatilhado, permanece em loop

        arquivo = open(completePath, "a+")
        rc = None
        while rc != win32event.WAIT_OBJECT_0:
            #Escrever algo no log
            now = datetime.datetime.now()
            arquivo.write(str(now)+" - Log de teste\n")
            arquivo.flush()
            #Espera 5 segundos para escrever a nova informacao no log
            rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)
         
        arquivo.write(str(datetime.datetime.now())+" - Fechando servico\n")
        arquivo.close()
    
    # chamado quando esta sendo desligado  
    def SvcStop(self):
        # diz ao scm que esta sendo desligado
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # inicia o evento de parada
        win32event.SetEvent(self.hWaitStop)
        
if __name__ == '__main__':
    win32serviceutil.HandleCommandLine(ServiceChecker)