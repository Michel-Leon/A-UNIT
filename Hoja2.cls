'=======================================================================================

Public Sub Worksheet_Activate()
    Bloqueo
    WITH Me.OleObjects("TipoMDPlano").object
        .clear
        .AddItem "PM Totalizador Interno"
        .AddItem "PM Totalizador Externo"
        .AddItem "PM Carga Interno"
        .AddItem "PM Carga Externo"
    End With
    Me.OleObjects("N_1").object.Value = ""
    Me.OleObjects("N_2").object.Value = ""
    Me.OleObjects("N_3").object.Value = ""
    Me.OleObjects("N_4").object.Value = ""
    Me.OleObjects("N_5").object.Value = ""
    Me.OleObjects("N_6").object.Value = ""
    Me.OleObjects("N_7").object.Value = ""
    Me.OleObjects("AWG_6").object.Value = ""
    Me.OleObjects("C_1").object.Value = ""
    Me.OleObjects("C_2").object.Value = ""
    Me.OleObjects("C_3").object.Value = ""
    Me.OleObjects("C_4").object.Value = ""
    Me.OleObjects("C_5").object.Value = ""
    Me.OleObjects("C_6").object.Value = ""
    Me.OleObjects("C_7").object.Value = ""
    Me.OleObjects("AWG_4").object.Value = ""
    Me.OleObjects("QT_1").object.Value = ""
    Me.OleObjects("CT_1").object.Value = ""
    Me.OleObjects("CT_2").object.Value = ""
    Me.OleObjects("CT_3").object.Value = ""
    Me.OleObjects("TAG_BARRA").object.Value = ""
    Me.OleObjects("Cor_N").object.Value = ""
    Me.OleObjects("Cor_C").object.Value = ""
    Me.OleObjects("PM_1").object.Value = ""
    Me.OleObjects("viene").object.Value = ""
End sub
Public sub Bloqueo()
    Me.OleObjects("N_1").object.Enabled = False
    Me.OleObjects("N_2").object.Enabled = False
    Me.OleObjects("N_3").object.Enabled = False
    Me.OleObjects("N_4").object.Enabled = False
    Me.OleObjects("N_5").object.Enabled = False
    Me.OleObjects("N_6").object.Enabled = False
    Me.OleObjects("N_7").object.Enabled = False
    Me.OleObjects("AWG_6").object.Enabled = False
    Me.OleObjects("C_1").object.Enabled = False
    Me.OleObjects("C_2").object.Enabled = False
    Me.OleObjects("C_3").object.Enabled = False
    Me.OleObjects("C_4").object.Enabled = False
    Me.OleObjects("C_5").object.Enabled = False
    Me.OleObjects("C_6").object.Enabled = False
    Me.OleObjects("C_7").object.Enabled = False
    Me.OleObjects("AWG_4").object.Enabled = False
    Me.OleObjects("QT_1").object.Enabled = False
    Me.OleObjects("CT_1").object.Enabled = False
    Me.OleObjects("CT_2").object.Enabled = False
    Me.OleObjects("CT_3").object.Enabled = False
    Me.OleObjects("TAG_BARRA").object.Enabled = False
    Me.OleObjects("Cor_N").object.Enabled = False
    Me.OleObjects("Cor_C").object.Enabled = False
    Me.OleObjects("PM_1").object.Enabled = False
    Me.OleObjects("viene").object.Enabled = False
end sub
Public sub Desbloqueo()
    Me.OleObjects("N_1").object.Enabled = True
    Me.OleObjects("N_2").object.Enabled = True
    Me.OleObjects("N_3").object.Enabled = True
    Me.OleObjects("N_4").object.Enabled = True
    Me.OleObjects("N_5").object.Enabled = True
    Me.OleObjects("N_6").object.Enabled = True
    Me.OleObjects("N_7").object.Enabled = True
    Me.OleObjects("AWG_4").object.Enabled = True
    Me.OleObjects("C_1").object.Enabled = True
    Me.OleObjects("C_2").object.Enabled = True
    Me.OleObjects("C_3").object.Enabled = True
    Me.OleObjects("C_4").object.Enabled = True
    Me.OleObjects("C_5").object.Enabled = True
    Me.OleObjects("C_6").object.Enabled = True
    Me.OleObjects("C_7").object.Enabled = True
    Me.OleObjects("AWG_6").object.Enabled = True
    Me.OleObjects("QT_1").object.Enabled = True    
    Me.OleObjects("CT_1").object.Enabled = True
    Me.OleObjects("CT_2").object.Enabled = True
    Me.OleObjects("CT_3").object.Enabled = True
    Me.OleObjects("TAG_BARRA").object.Enabled = True
    Me.OleObjects("Cor_N").object.Enabled = True
    Me.OleObjects("Cor_C").object.Enabled = True
    Me.OleObjects("PM_1").object.Enabled = True
    Me.OleObjects("viene").object.Enabled = True
    TAG
end sub
PUblic sub TAG()
    WITH Me.OleObjects("TAG_BARRA").object
        .clear
        .AddItem "VBB"
        .AddItem "HBB"
    End With
End sub

Public sub Medidores()
    '----------------------------------------------
        celda = "CC43"
        Set wsActual = thisWorkbook.Sheets("MD")
        'limpiar area
        'descombinar area
        thisWorkbook.Sheets("MD").Range("CC44:IQ147").UnMerge
        thisWorkbook.Sheets("MD").Range("CC44:IQ147").Clearcontents
    '----------------------------------------------
    dim dato as string
    dato = Me.OleObjects("TipoMDPlano").object
    Desbloqueo
    select case dato
        case "PM Totalizador Interno"
            Codigoresultado = "PLA_080601"
        case "PM Totalizador Externo"
            Codigoresultado = "PLA_080602"
        case "PM Carga Interno"
            Codigoresultado = "PLA_080603"
        case "PM Carga Externo"
            Codigoresultado = "PLA_080604"
    end select
    MoverBloque Codigoresultado,"BD_MEDIDOR", wsActual, celda
End sub
public sub comp_medidores()
    dim wsDestino as worksheet
    set wsDestino = ThisWorkbook.Worksheets("COMP_Medidores")
    dim dato1 as string, dato2 as string, dato3 as string, dato4 as string
    dim dato5 as string, dato6 as string, dato7 as string, dato8 as string
    dim dato9 as string, dato10 as string, dato11 as string, dato12 as string
    dim dato13 as string, dato14 as string, dato15 as string, dato16 as string
    dim dato17 as string, dato18 as string, dato19 as string, dato20 as string
    dim dato21 as string, dato22 as string, dato23 as string, dato24 as string
    dim dato25 as string
    '===========================================================================
    'RECOLEXION DE DATOS
    '===========================================================================
    'ACOMETIDA    
        dato1 = Me.OleObjects("N_1").object
        dato2 = Me.OleObjects("C_1").object
    'VBB
        dato3 = Me.OleObjects("N_2").object
        dato4 = Me.OleObjects("C_2").object
        dato18 = Me.OleObjects("TAG_BARRA").object
    'TAG INTERUPTOR TOTALIZADOR
        dato5 = Me.OleObjects("N_3").object
        dato6 = Me.OleObjects("C_3").object
        dato17 = Me.OleObjects("QT_1").object
    'SEÑALES DE TENSION
        dato7 = Me.OleObjects("N_4").object
        dato8 = Me.OleObjects("C_4").object
        dato15 = Me.OleObjects("AWG_4").object
    'SEÑALES DE NEUTRO
        dato9 = Me.OleObjects("N_5").object
        dato10 = Me.OleObjects("C_5").object
    'SEÑALES DE CORRIENTE
        dato11 = Me.OleObjects("N_6").object
        dato12 = Me.OleObjects("C_6").object
        dato16 = Me.OleObjects("AWG_6").object
        dato19 = Me.OleObjects("CT_1").object
        dato20 = Me.OleObjects("CT_2").object
        dato21 = Me.OleObjects("CT_3").object
    'ALIMENTACION
        dato13 = Me.OleObjects("N_7").object
        dato14 = Me.OleObjects("C_7").object
        dato25 = Me.OleObjects("viene").object
    'CORDENADS EN EL PLANO/ TAG PM    
        dato22 = Me.OleObjects("PM_1").object
        dato23 = Me.OleObjects("Cor_N").object
        dato24 = Me.OleObjects("Cor_C").object
    '===========================================================================
    'ESCRITURA DE DATOS
    '===========================================================================
    'CORDENADAS DEL PLANO
        wsDestino.Range("C39").Value = dato23
        wsDestino.Range("C40").Value = dato24
    'MEDIDOR
        wsDestino.Range("C7").Value = dato22
        wsDestino.Range("C8").Value = dato22
        wsDestino.Range("C10").Value = dato22
        wsDestino.Range("C31").Value = dato22
    'ACOMETIDA
        wsDestino.Range("C17").Value = dato1
        wsDestino.Range("C18").Value = dato2    
    'VBB   
        wsDestino.Range("C16").Value = dato18 
        wsDestino.Range("C19").Value = dato3
        wsDestino.Range("C20").Value = dato4    
    'TAG INTERUPTOR TOTALIZADOR
        wsDestino.Range("C6").Value = dato17
        wsDestino.Range("C11").Value = dato17
        wsDestino.Range("C25").Value = dato17
        wsDestino.Range("C28").Value = dato17
        'CORDENADAS INTERRUPTOR TOTALIZADOR
            wsDestino.Range("C21").Value = dato5
            wsDestino.Range("C22").Value = dato6
    'SEÑALES DE TENSION
        wsDestino.Range("C37").Value = dato15
        wsDestino.Range("C26").Value = dato7
        wsDestino.Range("C27").Value = dato8
    'SEÑALES DE NEUTRO
        wsDestino.Range("C29").Value = dato9
        wsDestino.Range("C30").Value = dato10 
    'SEÑALES DE CORRIENTE
        wsDestino.Range("C23").Value = dato11 
        wsDestino.Range("C24").Value = dato12  
        WsDestino.Range("C38").Value = dato16  
    'CT'S
        wsDestino.Range("C2").Value = dato19
        wsDestino.Range("C3").Value = dato20
        wsDestino.Range("C4").Value = dato21
        wsDestino.Range("C12").Value = dato19
        wsDestino.Range("C13").Value = dato20
        wsDestino.Range("C14").Value = dato21   
    'ALIMENTACION
        wsDestino.Range("C35").Value = dato13
        wsDestino.Range("C36").Value = dato14 
        wsDestino.Range("C32").Value = dato25   
END SUB    


