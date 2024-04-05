// Programa   : DPIMPRXLSDEF
// Fecha/Hora : 06/03/2014 01:29:13
// Propósito  : Definición de Importación de Archivos
// Creado Por : Juan Navas
// Llamado por: DPIMPRXLS
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(cCodigo,lRun)
  LOCAL aRef:={},cTable:="",nAt,cTabla,cFileXls,bPostSave,cMemo:="",cDescri:=""
  LOCAL aTablas:={},cFileIxl

/*
  AADD(aTablas,{"Productos"         ,"DPINV"      })
  AADD(aTablas,{"Proveedores"       ,"DPPROVEEDOR"})
  AADD(aTablas,{"Clientes"          ,"DPCLIENTES" })
  AADD(aTablas,{"Cuentas por Pagar" ,"DPDOCPRO"   })
  AADD(aTablas,{"Cuentas por Cobrar","DPDOCCLI"   })
  AADD(aTablas,{"Plan de Cuentas"   ,"DPCTA"      })
  AADD(aTablas,{"Libro de Compras"  ,"DPDOCPRO"   })
  AADD(aTablas,{"Activos"           ,"DPACTIVOS"  })

  AEVAL(aTablas,{|a,n| SQLUPDATE("DPIMPRXLS","IXL_TABLA",a[2],"IXL_PREDEF"+GetWhere("=",a[1]))})
*/

//  OpenTable("SELECT * FROM   DPIMPRXLS",.F.):oOdbc:Execute("UPDATE DPIMPRXLS SET IXL_TABLE=IXL_TABLA WHERE IXL_TABLE IS NULL") 

  DEFAULT cCodigo:=SQLGET("DPIMPRXLS","IXL_CODIGO")

  cTable  :=ALLTRIM(SQLGET("DPIMPRXLS","IXL_TABLA,IXL_FILE,IXL_MEMO,IXL_DESCRI","IXL_CODIGO"+GetWhere("=",cCodigo)))

  IF Empty(cTable)
     MsgMemo("Registro "+cCodigo+" no tiene Registro de Tabla")
     RETURN .F.
  ENDIF

  cFileXls:=IF(Empty(oDp:aRow),"",oDp:aRow[2])
  cMemo   :=IF(Empty(oDp:aRow),"",oDp:aRow[3])
  cDescri :=IF(Empty(oDp:aRow),"",oDp:aRow[4])

  IF Empty(cFileXls)
     MensajeErr("Es necesario Indicar Archivo Excel")
     RETURN .F.
  ENDIF

//cTabla:=cTable

/*
  nAt     :=ASCAN(aTablas,{|a,n| a[1]=cTable })

  IF nAt=0 .AND. !Empty(SQLGET("DPTABLAS","TAB_NOMBRE,TAB_DESCRI","TAB_NOMBRE"+GetWhere("=",cTable)))
     EJECUTAR("IXLCREATE",cTable,cDescri,cFileIxl,aRef,bPostSave,NIL,cFileXls,cCodigo) 
     RETURN .T.
  ENDIF

  IF nAt=0
     MensajeErr("Tabla "+cTable+" no definida")
     RETURN .F.
  ENDIF

//  cTabla:=aTablas[nAt,2]
*/

/*

  IF cTable="DPINV"
    AADD(aRef,{"@BARRA"    ,"Código de Barra"})
    AADD(aRef,{"@PRECIO_A" ,"Precio A"})
    AADD(aRef,{"@PRECIO_B" ,"Precio B"})
    AADD(aRef,{"@PRECIO_C" ,"Precio C"})
    AADD(aRef,{"@PRECIO_D" ,"Precio D"})
    AADD(aRef,{"@PRECIO_E" ,"Precio E"})
    AADD(aRef,{"@UNDMED"   ,"Unidad de Medida"})
    AADD(aRef,{"@CXUNDMED" ,"Cantidad por Unidad de Medida"})
    AADD(aRef,{"@PESO"     ,"Peso por Unidad de Medida"})
    AADD(aRef,{"@PRESENTA" ,"Presentación en Unidad medida"})
    AADD(aRef,{"@CANT"     ,"Cantidad Existencia"})
    AADD(aRef,{"@COSTO"    ,"Costo segun cada Fracción de Cantidad"})
    AADD(aRef,{"@LOTE"     ,"Numero de Lote"})
    AADD(aRef,{"@FCHVENC"  ,"Fecha de Vencimiento"})
    AADD(aRef,{"@PRECIO_L" ,"Precio Lote"})
    AADD(aRef,{"@CODSUC"   ,"Código Sucursal"})
    AADD(aRef,{"@CODALM"   ,"Almacen para la Existencia indicada en @CANT"})
    AADD(aRef,{"@LIMSUC"   ,"Limitar Sucursal SI o NO "}) 
    AADD(aRef,{"@GRUNOMBRE","Nombre de "+oDp:xDPGRU})
    AADD(aRef,{"@MARNOMBRE","Nombre de "+oDp:xDPMARCAS})
  ENDIF

  IF cTable="DPDOCPRO"
    AADD(aRef,{"@PRO_NOMBRE"    ,"Nombre del Proveedor"})
    AADD(aRef,{"@PRO_RIF"       ,"Numero del Rif"})
  ENDIF

  IF cTable="DPDOCCLI"
    AADD(aRef,{"@CLI_NOMBRE"    ,"Nombre del Cliente"})
    AADD(aRef,{"@CLI_RIF"       ,"Numero del Rif"})
//    AADD(aRef,{"@CLI_CODSUC"    ,"Código de Sucursal"})
  ENDIF

  IF cTable="DPACTIVOS"
    AADD(aRef,{"@MTODEP"   ,"Monto Depreciación"})
    AADD(aRef,{"@GRUNOMBRE","Nombre de "+oDp:xDPGRUACTIVOS})
    AADD(aRef,{"@UBINOMBRE","Nombre de "+oDp:xDPUBIACTIVOS})
  ENDIF

//  IF cTable="DPASIENTOS"
//    aRef:=EJECUTAR("DPIMPRXLSASIENTOSXLS")
//    AADD(aRef,{"@CTANOMBRE","Nombre de "+oDp:xDPCTA})
//    AADD(aRef,{"@CBTNOMBRE","Nombre de Comprobante"})
//    AADD(aRef,{"@DEBE"     ,"Monto Debe "})
//    AADD(aRef,{"@HABER"    ,"Monto Haber"})
//  ENDIF
 
  IF cTable="NMTRABAJADOR"
    AADD(aRef,{"@DEP_DESCRI","Descripción de "+oDp:xDPDPTO})
    AADD(aRef,{"@GTR_DESCRI","Descripción de "+oDp:xNMGRUPO})
    AADD(aRef,{"@CAR_DESCRI","Descripción de "+oDp:xNMCARGOS})
  ENDIF
*/

  EJECUTAR("IXLCREATE",cTable,cDescri,cFileIxl,aRef,bPostSave,NIL,cFileXls,cCodigo) 

RETURN .T.
// EOF
