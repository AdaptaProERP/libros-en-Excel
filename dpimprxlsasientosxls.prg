// Programa   : DPIMPRXLSASIENTOSXLS
// Fecha/Hora : 22/11/2022 20:51:22
// Propósito  : Importar Asientos desde excel
// Creado Por :
// Llamado por:
// Aplicación :
// Tabla      :

#INCLUDE "DPXBASE.CH"

PROCE MAIN(lGetRef)
  LOCAL oDpLbx,aRef:={}

  DEFAULT lGetRef:=.F.

  AADD(aRef,{"@CTANOMBRE","Nombre de "+oDp:xDPCTA})
  AADD(aRef,{"@CBTNOMBRE","Nombre de Comprobante"})
  AADD(aRef,{"@DEBE"     ,"Monto Debe "})
  AADD(aRef,{"@HABER"    ,"Monto Haber"})

  AEVAL(aRef,{|a,n| AADD(aRef[n],"C")})

  IF lGetRef
    RETURN aRef
  ENDIF

  oDpLbx:=TDpLbx():New("DPIMPRXLS.LBX",NIL,[IXL_TABLA="DPASIENTOS"],NIL,NIL,{"DPASIENTOS","Asientos",ACLONE(aRef)})
  oDpLbx:uValue1:="DPCLIENTES"
  oDpLbx:uValue2:="Clientes"
  oDpLbx:uValue3:=ACLONE(aRef)
  oDpLbx:Activate()

RETURN aRef
// EOF

