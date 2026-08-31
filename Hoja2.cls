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
End sub
Public sub Bloqueo()
    Me.OleObjects("N_1").object.Enabled = False
    Me.OleObjects("N_2").object.Enabled = False
    Me.OleObjects("N_3").object.Enabled = False
    Me.OleObjects("N_4").object.Enabled = False
    Me.OleObjects("N_5").object.Enabled = False
    Me.OleObjects("N_6").object.Enabled = False
    Me.OleObjects("AWG_6").object.Enabled = False
    Me.OleObjects("N_8").object.Enabled = False
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
    dim dato21 as string, dato22 as string
    '===========================================================================
    'RECOLEXION DE DATOS
    '===========================================================================
    dato1 = Me.OleObjects("N_1").object
    dato2 = Me.OleObjects("C_1").object
    dato3 = Me.OleObjects("N_2").object
    dato4 = Me.OleObjects("C_2").object
    dato5 = Me.OleObjects("N_3").object
    dato6 = Me.OleObjects("C_3").object
    dato7 = Me.OleObjects("N_4").object
    dato8 = Me.OleObjects("C_4").object
    dato15 = Me.OleObjects("AWG_4").object
    dato9 = Me.OleObjects("N_5").object
    dato10 = Me.OleObjects("C_5").object
    dato11 = Me.OleObjects("N_6").object
    dato12 = Me.OleObjects("C_6").object
    dato16 = Me.OleObjects("AWG_6").object
    dato13 = Me.OleObjects("N_7").object
    dato14 = Me.OleObjects("C_7").object
    dato17 = Me.OleObjects("QT_1").object
    dato18 = Me.OleObjects("TAG_BARRA").object
    dato19 = Me.OleObjects("CT_1").object
    dato20 = Me.OleObjects("CT_2").object
    dato21 = Me.OleObjects("CT_3").object
    dato22 = Me.OleObjects("PM_1").object
    '===========================================================================
    'ESCRITURA DE DATOS
    '===========================================================================
    'CT'S
        wsDestino.Range("C2").Value = dato19
        wsDestino.Range("C3").Value = dato20
        wsDestino.Range("C4").Value = dato21
        wsDestino.Range("C16").Value = dato19
        wsDestino.Range("C17").Value = dato20
        wsDestino.Range("C18").Value = dato21
    ' TAG INTERUPTOR TOTALIZADOR
        wsDestino.Range("C6").Value = dato17
        wsDestino.Range("C11").Value = dato17
        wsDestino.Range("C25").Value = dato17
        wsDestino.Range("C28").Value = dato17
        ' CORDENADAS INTERRUPTOR TOTALIZADOR
            wsDestino.Range("C21").Value = dato5
            wsDestino.Range("C22").Value = dato6
    'MEDIDOR
        wsDestino.Range("C7").Value = dato22
        wsDestino.Range("C8").Value = dato22
        wsDestino.Range("C10").Value = dato22
        wsDestino.Range("C31").Value = dato22
    'ACOMETIDA
        wsDestino.Range("C17").Value = dato1
        wsDestino.Range("C18").Value = dato2
    ' VBB   
        wsDestino.Range("C16").Value = dato18 
        wsDestino.Range("C19").Value = dato3
        wsDestino.Range("C20").Value = dato4
END SUB    


