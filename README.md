# AccessUpdaterVBAExcel
Actualizador de base de datos access usando VBA Excel

1. La clase llamada "cls_RECUPERA_ESTRUCTURA_BDLOCAL" analiza una base de datos Microsoft Access y guarda su estructura en un diccionario, además permite generar el código VBA Excel necesario para replicar esa misma estructura.
2.La clase llamada "cls_ACTUALIZADOR_BD_LOCAL" utiliza a "cls_RECUPERA_ESTRUCTURA_BDLOCAL" para recuperar la estructura en forma de diccionario de una base de datos Microsoft Access que será igualada en estructura a otra base de datos Microsoft Access guardada en forma de clase con los metodos necesarios para su creación.
