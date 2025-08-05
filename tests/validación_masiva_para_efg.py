import modalidad_40.calculos.mejora_pension as mp 
import sys, os,json, datetime
import pandas as pd

ruta_archivo=sys.argv[1]
carpeta_output=sys.argv[2]
fechastring=datetime.datetime.now().strftime('%Y%m%d')
ruta_output=carpeta_output+"\\"+fechastring+' Output M40.xlsx'


def genr_nombrearchivos(ruta):
    files=[]
    for (dirpath, dirnames, filenames) in os.walk(ruta):
        files.extend(filenames)
        break
    files=[ruta+'\\'+i for i in files]     
    return files

#Leemos rutas de archivos
archivos=genr_nombrearchivos(ruta_archivo)

#Generamos Resultados
resultados={}


with pd.ExcelWriter(ruta_output) as writer:
    rechazos=pd.Series()
    for ruta_sisec in archivos:
        sisec=mp.SISECpdf(ruta_sisec,mp.diccionario_entidades_federativas)
        print(f'Comenzando {sisec.datos_cliente.Nombre}')
        análisis_semanas=mp.Análisis_Semanas(sisec)
        diagnóstico=mp.Diagnóstico_Mejora_Pensión(sisec,análisis_semanas,obtener_json=False)
        if diagnóstico.diagnóstico:
            print('Diagnóstico Ok')
            #Propuesta
            propuesta=mp.Carpeta_M40(sisec.datos_cliente.Nombre,sisec.datos_cliente.Edad,sisec.datos_cliente.nss,sisec.datos_cliente.curp,análisis_semanas.salario_promedio,\
                                sisec.datos_cliente.semanas_cotizadas,sisec.datos_cliente.fecha_emisión,análisis_semanas.fecha_baja,análisis_semanas.fecha_inicio_laboral,análisis_semanas.ultimo_salario_mensual\
                                    ,0,0,0,0,mp.cuantía_incrementos,mp.tablas_crédito,elaborar_json=False)    
            propuesta.tabla_resumen=propuesta.tabla_resumen[['nombre','tipo_calculo','edad','semanas_cotizadas','fecha_baja_ultimo_patron_real','salario_promedio','pension_neta_mensual','ultimo_salario_mensual','fecha_inicio_laboral','cobro_80','aportacion_efg','subtotal','costo_proyecto']].copy()
            serie_original=propuesta.tabla_resumen.loc['original'].dropna()
            serie_original.index=pd.MultiIndex.from_tuples([('original',i) for i in serie_original.index])

            serie_mejorada=propuesta.tabla_resumen.loc['mejorada'].dropna()
            serie_mejorada.index=pd.MultiIndex.from_tuples([('mejorada',i) for i in serie_mejorada.index])

            serie_incrementos=propuesta.tabla_resumen.loc['incrementos'].dropna()
            serie_incrementos.index=pd.MultiIndex.from_tuples([('inc8807rementos',i) for i in serie_incrementos.index])

            serie_financiamiento=propuesta.tabla_resumen.loc['financiamiento']
            serie_financiamiento.index=pd.MultiIndex.from_tuples([('financiamiento',i) for i in serie_financiamiento.index]) 

            ##Modificamos los datos para que tengan un nivel adicional
            
            # datos.index=pd.MultiIndex.from_tuples([('Input',i) for i in datos.index])

            # nueva_serie=pd.concat([datos,serie_original.dropna(),serie_mejorada.dropna(),serie_incrementos.dropna(),serie_financiamiento.dropna()])
            nueva_serie=pd.concat([serie_original.dropna(),serie_mejorada.dropna(),serie_incrementos.dropna(),serie_financiamiento.dropna()])
            nuevo_df=pd.DataFrame(nueva_serie)
            try:
                nuevo_df.T.to_excel(writer,sheet_name=sisec.datos_cliente.Nombre[0:10])

                tabla_amortizacion=propuesta.tabla_amortizacion.copy()
                tabla_amortizacion['fecha_baja']=propuesta.fecha_baja
                tabla_amortizacion['fecha_baja_reactivación']=propuesta.fecha_baja_reactivación
                # tabla_amortizacion['meses_mod_40_ajustado']=propuesta.meses_mod_40_ajustado
                # tabla_amortizacion['valor_uma']=propuesta.uma
                tabla_amortizacion['salario_a_contratar']=propuesta.salario_mensual_a_contratar
                tabla_amortizacion['salario_promedio_nuevo']=propuesta.salario_promedio_nuevo
                tabla_amortizacion.to_excel(writer,sheet_name='TA_'+sisec.datos_cliente.Nombre[0:10])
                print('Exportado')
            except:
                print('No se generan datos. Revisar SISEC')


            # dataframe_final=pd.concat([dataframe_final,nuevo_df])

            # nuevo_df.T.to_excel(rf'Prueba {nombre[0:10]}.xlsx')
        else:
            rechazos[sisec.datos_cliente.Nombre]=diagnóstico.mensaje
            print('Diagnóstico rechazo: Consultar pestaña rechazos')

            pass
    df_rechazos=pd.DataFrame(rechazos,columns=['Mensaje Rechazo'])
    df_rechazos.to_excel(writer,sheet_name='Rechazos')
    print(os.getcwd())              



#Generamos Dataframe



# dataframe.to_excel('Validación Cálculos.xlsx',merge_cells=False)
