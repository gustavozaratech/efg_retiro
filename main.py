import modalidad_40.calculos.mejora_pension as mp 
import sys

ruta_sisec=sys.argv[1]
sar_97,sar_92,infonavit_sar_97,infonavit_sar_92=float(sys.argv[2]),float(sys.argv[3]),float(sys.argv[4]),float(sys.argv[5])


sisec=mp.SISECpdf(ruta_sisec,mp.diccionario_entidades_federativas)
análisis_semanas=mp.Análisis_Semanas(sisec)
diagnóstico=mp.Diagnóstico_Mejora_Pensión(sisec,análisis_semanas,obtener_json=False)

#Condicionamos diagnóstico

if diagnóstico.diagnóstico:
    #Propuesta
    carpeta=mp.Carpeta_M40(sisec.datos_cliente.Nombre,sisec.datos_cliente.Edad,sisec.datos_cliente.nss,sisec.datos_cliente.curp,análisis_semanas.salario_promedio,\
                           sisec.datos_cliente.semanas_cotizadas,sisec.datos_cliente.fecha_emisión,análisis_semanas.fecha_baja,análisis_semanas.fecha_inicio_laboral,análisis_semanas.ultimo_salario_mensual,sar_97,sar_92,infonavit_sar_97,infonavit_sar_92,mp.cuantía_incrementos,mp.tablas_crédito)
    
    # carpeta.tabla_resumen.to_excel('D:\OneDrive\Documentos\ANALYTICS PYMES\CLIENTES\EFGMEXICO\CALCULADORA_PENSION\pruebas'+f'\\{sisec.datos_cliente.Nombre}.xlsx')
else:
    pass    



