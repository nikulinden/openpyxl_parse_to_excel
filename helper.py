import os

copying_from_game = r'C:\PROJECTS\SJPD-Wild_Panda\perforce\WildPanda\Targets\Math\Game_XML\Variations'
copying_from_common = r'C:\PROJECTS\SJPD-Wild_Panda\perforce\Common\Targets\Math\Game_XML\Variation'

dir = os.getcwd()
#dir = 'r\''+str(script)+'\''
for filename in os.listdir(dir):
    if filename.startswith("wildpanda"):
        filename = filename.replace('wildpanda_','')
