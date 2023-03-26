@echo Aguarde backup em curso ... n�o feche esta janela
@sqlcmd -S "localhost\primavera" -U sa -P ola2022. -i "C:\SQL\BACKUPScript\primavera.sql" >backup.log
@echo Backup terminado

@echo Opera��o terminada. Pode fechar esta janela. Obrigado.