# -*- coding: utf-8 -*-
"""
Created on Fri Aug 13 11:31:34 2021

@author: carlo
"""


import pymysql
#from tkinter import ttk

def Conectar():
    db = pymysql.connect(
        host='localhost',
        user='root', 
        password = 'master1090',
        db='contabilidad',
        )
    
    return db


def Ejecutar(db,sql):
    
    cursor = db.cursor()
    try:
       # Execute the SQL command
       cursor.execute(sql)
       # Commit your changes in the database
       db.commit()
    except:
       # Rollback in case there is any error
       db.rollback()
       print("error en la base de datos")


def Desconectar(db):
    db.close()
    
def ElegirTipo(tipoOpcion):
    if tipoOpcion == 1:
        return "activo"
    if tipoOpcion == 2:
        return "pasivo"
    if tipoOpcion == 3:
        return "patrimonio"
    if tipoOpcion == 4:
        return "gastos"
    if tipoOpcion == 5:
        return "utilidades"

def CrearCuenta(rut,nombre,codigo,tipo,ventana):
    
    db = Conectar()
    
    sql = "INSERT INTO cuenta(nombre, codigo, tipo,rut) VALUES ('{0}','{1}','{2}','{3}');".format(nombre,codigo,tipo,rut)
    Ejecutar(db,sql)
    
    Desconectar(db)
    ventana.destroy()
    
def EditarCuenta(nombre,cod,tipo,rut,newCod,ventana):
    
    db = Conectar()
    
    sql = "UPDATE cuenta SET tipo = '{0}', codigo = '{1}',nombre = '{2}' WHERE rut = '{3}' AND codigo = '{4}';".format(tipo,newCod,nombre,rut,cod)
    
    Ejecutar(db,sql)
    
    Desconectar(db)    
    
    ventana.destroy()
    
    
def MostrarCuentas(rut):
    db = Conectar()
    sql = "SELECT nombre, codigo, tipo FROM cuenta WHERE rut = '{0}';".format(rut)
    cursor = db.cursor()
    results = None
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
    except:
        db.rollback()
        print("error en la base de datos")
        return []
    
    return results
       
def BorrarCuenta(rut,cod,ventana):
    
    db = Conectar()
    
    sql = "DELETE FROM cuenta WHERE rut = '{0}' AND codigo = '{1}';".format(rut,cod)
    Ejecutar(db,sql)
    
    Desconectar(db)
    
    ventana.destroy()
    
def CrearEmpresa(nombre,duenyo,rut,rutProp, direccion, giro,ventana):
    db = Conectar()
    
    sql = "INSERT INTO empresa(nombre, dueño, rut, rutProp, direccion, giro) VALUES ('{0}','{1}','{2}','{3}','{4}','{5}');".format(nombre,duenyo,rut,rutProp, direccion, giro)
   
    Ejecutar(db,sql)
    
    Desconectar(db)
    
    ventana.destroy()
    
def MostrarEmpresas():
    db = Conectar()
    
    cursor = db.cursor()
    
    sql = "SELECT nombre, rut, dueño, rutProp, direccion, giro FROM empresa;"
    results = None
    try:
        # Execute the SQL command
        cursor.execute(sql)
        
        results = cursor.fetchall()
    except:
        print("error en la base de datos")
        db.rollback()
        Desconectar(db)
        return []
    
    Desconectar(db)
    return results
    
    
    
def SeleccionarEmpresa():
    
    ruts = MostrarEmpresas()
    
    index = int(input("seleccione la empresa\n"))
    
    return ruts[index - 1]

def BorrarEmpresa(rut):
    db = Conectar()
    
    sql = "DELETE FROM empresa WHERE rut = '{0}';".format(rut)
    Ejecutar(db,sql)
    
    Desconectar(db)
    
def EditarEmpresa(nombre,duenyo,rut,newRut,rutProp,direccion,giro,ventana):
    #ruts = MostrarEmpresas()

    db = Conectar()
    sql = "UPDATE empresa SET nombre = '{0}', dueño = '{1}', rut = '{2}' ,rutProp = '{4}',direccion = '{5}',giro = '{6}' WHERE rut = '{3}';".format(nombre,duenyo,newRut,rut,rutProp,direccion,giro)
    Ejecutar(db,sql)
    
    Desconectar(db)
    
    ventana.destroy()

def CrearTransaccion(cuentas,nombresCuentas):
    debe = input("ingrese el debe")
    haber = input("ingrese el haber")
    
    print("elija la cuenta a ingresar:")
    for cuenta in nombresCuentas:
        print(cuenta)
        
    nombreCuenta = input()
    
    return (debe,haber,cuentas[nombreCuenta])

def ObtenerComprobantes(rut):
    db = Conectar()
    
    cursor = db.cursor()
    
    sql = "SELECT codigoCom, glosa, tipo, fecha FROM comprobante WHERE rut = '{0}'".format(rut)
    results = None
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
        
    except:
        print("error en la base")
        db.rollback()
        return
    
    return results

def MostrarComprobantes():
    db = Conectar()
    
    cursor = db.cursor()
    
    
    sql = "SELECT codigoCom,fecha,tipo,glosa FROM comprobante"
    
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
        for row in results:
            print("codigoCom: {0}".format(row[0]))
            print("fecha: {0}".format(row[1]))
            print("tipo: {0}".format(row[2]))
            print("glosa: {0}".format(row[3]))
    except:
        print("error en la base")
        db.rollback()
        
    Desconectar(db)

def MostrarTransaccion(listaTransacciones):
    
    db = Conectar()
    cursor = db.cursor()
    i = 1
    for item in listaTransacciones:
        print("transaccion: {0}".format(i))
        print("debe: '{0}'".format(item[0]))
        print("haber: '{0}'".format(item[1]))
        i += 1
        sql = "SELECT nombre,codigo FROM cuenta WHERE codigo = '{0}';".format(item[2])
        
        try:
            cursor.execute(sql)
            results2 = cursor.fetchall()
            
            for row2 in results2:
                print("nombre cuenta: '{0}'".format(row2[0]))
                print("codigo: '{0}'".format(row2[1]))
        except:
            print("error en la base")
            db.rollback()
            Desconectar(db)
            return
            
    Desconectar(db)
    
    
def ObtenerTransacciones(rut,codigoCom):
    db = Conectar()
    cursor = db.cursor()
    
    sql = "SELECT debe,haber, c.codigo, nombre, tipo FROM transaccion t inner join cuenta c on t.codigo = c.codigo and t.rut = c.rut where t.codigoCom = '{0}' and t.rut = '{1}';".format(codigoCom,rut)
    
    resultado = None
    
    try:
        cursor.execute(sql)
        resultado = cursor.fetchall()
   
    except:
        print("error en la base")
        db.rollback()
        Desconectar(db)
        return
    
    return resultado
    
    

def BorrarTransaccion(listaTransacciones):
    MostrarTransaccion(listaTransacciones)
    
    index = int(input("seleccione la transaccion a borrar, u otro boton para salir"))
    
    if index > len(listaTransacciones):
        return
    
    borrar = listaTransacciones[index - 1]    
    listaTransacciones.remove(borrar)
    
def EditarTransaccion(listaTransacciones,cuentas,nombresCuentas):
    MostrarTransaccion(listaTransacciones)
    
    index = int(input("seleccione la transaccion a modificar, u otro boton para salir"))
    
    if index > len(listaTransacciones):
        return
    
    opcion = input("1: cambiar debe\n2:cambiar haber\n3: cambiar cuenta,u otro boton para salir")
    
    tuplaAux = None
    if opcion == 1:
        debe = input("ingrese el nuevo debe")
        
        tuplaAux = (debe,listaTransacciones[index][1],listaTransacciones[index][2])
    
    elif opcion == 2:
        haber = input("ingrese el nuevo haber")
        
        tuplaAux = (listaTransacciones[index][0],haber,listaTransacciones[index][2])
        
    elif opcion == 3:
        print("elija la cuenta a ingresar:")
        for cuenta in nombresCuentas:
            print(cuenta)
            
        nombreCuenta = input()
        
        tuplaAux = (listaTransacciones[index][0],listaTransacciones[index][1],cuentas[nombreCuenta])
        
    else:
        return
    
    listaTransacciones[index] = tuplaAux
    

def DefinirCuentas(rut):
    db = Conectar()
    
    cursor = db.cursor()
    
    cuentas = {}
    
    nombresCuentas = []
    
    sql = "SELECT nombre, codigo FROM cuenta WHERE rut = '{0}';".format(rut)
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
        for row in results:
           nombresCuentas.append(row[0])
           cuentas[row[0]] = row[1]
    except:
        print("error en la base")
        db.rollback()
    
    Desconectar(db)
    
    return cuentas, nombresCuentas

def ComprobarDebeHaber(listaTransacciones):
    sumDebe = 0
    sumHaber = 0
    print("lista transacciones: " + str(listaTransacciones))
    for item in listaTransacciones:
        auxDebe, auxHaber, _ = item.ObtenerDatos()
        
        sumDebe += auxDebe
        sumHaber += auxHaber
    
    if sumDebe != sumHaber:
        print("su debe y haber no son iguales, intente de nuevo")
        return -1
    
    return 0

def IngresarComprobante(rut,fecha,tipo,glosa,listaTransacciones,ventana):
    print(listaTransacciones)
    if ComprobarDebeHaber(listaTransacciones) == -1:
        return 0
    
    db = Conectar()
    
    print("fecha: " + str(fecha))
    
    fechaAux = str(fecha)
    
    print("fechaAux: " + str(fechaAux))
    
    fechaAux2 = fechaAux.split("-")
    
    print("fechaAux: " + str(fechaAux2))
    
    fechaInicio = fechaAux2[0] + "-" + fechaAux2[1] + "-01"
    fechaFinal = fechaAux2[0] + "-" + fechaAux2[1] + "-31"
    
    print("fechaInicio: " + str(fechaInicio))
    print("fechaFinal: " + str(fechaFinal))
    
    sql = "SELECT COUNT(*) FROM comprobante WHERE fecha BETWEEN '{0}' AND '{1}'".format(fechaInicio,fechaFinal)
    print(sql)
    cursor = db.cursor()
    
    numMes = 0
    
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
        numMes = results[0][0]
    except:
        print("error en la base")
        db.rollback()
        Desconectar(db)
        return
    
    sql = "INSERT INTO comprobante(rut, fecha, tipo,glosa,numMes) VALUES ('{0}',STR_TO_DATE('{1}', '%Y-%m-%d'),'{2}','{3}',{4});".format(rut,fecha,tipo,glosa,numMes + 1)
    Ejecutar(db,sql)
    sql = "SELECT MAX(codigoCom) FROM comprobante;"
    cursor = db.cursor()
    codCodigo = 0
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
        codCodigo = results[0][0]
    except:
        print("error en la base")
        db.rollback()
        Desconectar(db)
        return
    
    #num = len(listaTransacciones)
    
    for item in listaTransacciones:
        auxDebe,auxHaber,auxCodCuenta = item.ObtenerDatos()
        
        sql = "INSERT INTO transaccion(rut, codigoCom, debe, haber, codigo) VALUES ('{0}','{1}','{2}','{3}','{4}');".format(rut,codCodigo,auxDebe,auxHaber,auxCodCuenta)
        Ejecutar(db,sql)

    Desconectar(db)
    
    print("comprobante ingresado")
    ventana.destroy()
    return 1

    
def CrearComprobante(rut):
    
    fecha = input("ingrese la fecha del comprobante\n")
    tipo = input("ingrese el tipo: I, E o T\n")
    glosa = input("ingrese glosa\n")
    
    cuentas, nombresCuentas = DefinirCuentas(rut)
    
    seguro = True
    
    listaTransacciones = []
    
    while(seguro):
        
        print("seleccione operacion:\n1: agregar transaccion\n2:borrar transaccion\n3:editar transaccion\n4: crear comprobante\n5: salir\n")
        opcion = int(input())
        
        if opcion == 1:
            
            auxTupla = CrearTransaccion(cuentas,nombresCuentas)
            if int(auxTupla[0]) == 0 and int(auxTupla[1]) == 0:
                print("por favor introdusca algun valor mayor a cero en el debe o el haber")
            else:
                listaTransacciones.append(auxTupla)
            
            
            
        elif opcion == 2:
            BorrarTransaccion(listaTransacciones)
        elif opcion == 3:
            EditarTransaccion(listaTransacciones,cuentas,nombresCuentas)
        elif opcion == 4:
            
            aux = IngresarComprobante(rut,fecha,tipo,glosa,listaTransacciones)
            
            if aux == -1:
                continue
            
            seguro = False
            
        elif opcion == 5:
            return            

    
def EditarComprobante(rut,codigoCom,fecha,tipo,glosa,listaTransacciones,ventana):
    if ComprobarDebeHaber(listaTransacciones) == -1:
        return 0
    
    db = Conectar()
    
    #cursor = db.cursor()
    fecha = None
    tipo = None
    
    sql = "UPDATE comprobante SET fecha = STR_TO_DATE('{0}', '%Y-%m-%d'), tipo = '{1}, glosa = '{2}' WHERE rut = '{3}' AND codigoCom = '{4}'".format(fecha,tipo,glosa,rut,codigoCom)
    
    Ejecutar(db,sql)
    
    sql = "DELETE FROM transaccion WHERE rut= '{0}' AND codigoCom = '{1}';".format(rut,codigoCom)
    Ejecutar(db,sql)
    
    
    for item in listaTransacciones:
        auxDebe,auxHaber,auxCodCuenta = item.ObtenerDatos()
        
        sql = "INSERT INTO transaccion(rut, codigoCom, debe, haber, codigo) VALUES ('{0}','{1}','{2}','{3}','{4}');".format(rut,codigoCom,auxDebe,auxHaber,auxCodCuenta)
        Ejecutar(db,sql)
    
    return 1
        
def BorrarComprobante(rut,codigoCom):      
    db = Conectar()
    
    sql = "DELETE FROM transaccion WHERE rut= '{0}' AND codigoCom = '{1}';".format(rut,codigoCom)
    Ejecutar(db,sql)
    
    sql = "DELETE FROM comprobante WHERE rut= '{0}' AND codigoCom = '{1}';".format(rut,codigoCom)
    Ejecutar(db,sql)
    
    Desconectar(db)
        
def VerMayor(codCuenta,rut,fechaInicial,fechaFinal):
    db = Conectar()
    
    cursor = db.cursor()
    
    sql = "SELECT nombre,tipo,codigo FROM cuenta WHERE ("
    
    for c in codCuenta:
        sql += " codigo = '{0}' OR".format(c)
        
    sql +=" codigo = 'aux') AND rut = '76.220.175-5';"

    nombre = None
    tipo = None
   
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
        
    except:
        print("error en la base1")
        db.rollback()
        Desconectar(db)
        return
    
    for row in results:
        nombre = row[0]
        tipo = row[1]
        totalDebe = 0
        totalHaber = 0
    
        #listaTransacciones = []
        sql = "SELECT debe,haber,codigoCom FROM transaccion AS T WHERE codigo = '{0}' AND rut = '{1}' AND (SELECT fecha FROM comprobante WHERE codigoCom = T.codigoCom) BETWEEN STR_TO_DATE('{2}', '%Y-%m-%d') AND STR_TO_DATE('{3}', '%Y-%m-%d') ORDER BY (SELECT fecha FROM comprobante WHERE codigoCom = T.codigoCom) ASC;".format(row[2],rut,fechaInicial,fechaFinal)
        try:
            print("nombre cuenta: '{0}'".format(nombre))
            print("tipo cuenta: '{0}'".format(tipo))
            cursor.execute(sql)
            results2 = cursor.fetchall()
            
            i = 1
            for row2 in results2:
                sql = "SELECT fecha,tipo,glosa FROM comprobante WHERE codigoCom = '{0}';".format(row2[2])
                try:
                    cursor.execute(sql)
                    results3 = cursor.fetchall()
                except:
                    print("error en la base2")
                    db.rollback()
                    Desconectar(db)
                    return
                
                if len(results3) == 0:
                    continue
                
                print("transaccion '{0}'".format(i))
                print("fecha: '{0}'".format(results3[0][0]))
                print("tipo: '{0}'".format(results3[0][1]))
                print("glosa: '{0}'".format(results3[0][2]))
                print("debe: {0}".format(row2[0]))
                print("haber: {0}".format(row2[1]))
                print("############################")
                totalDebe += int(row2[0])
                totalHaber += int(row2[1])
                i += 1
        except:
            print("error en la base3")
            db.rollback()
            Desconectar(db)
            return
        
        print("totalDebe: '{0}'".format(totalDebe))
        print("totalHaber: '{0}'".format(totalHaber))
    
            
    Desconectar(db)
    
    
def ObtenerDiario(rut,fechaInicial,fechaFinal):
    db = Conectar()
    
    cursor = db.cursor()
    
    sql = "SELECT c.fecha, c.numMes, c.glosa, t.debe, t.haber, cu.codigo, cu.nombre, c.tipo FROM comprobante c inner join transaccion t inner join cuenta cu on c.codigoCom = t.codigoCom AND t.codigo = cu.codigo AND c.rut = t.rut AND cu.rut = t.rut WHERE c.rut = '{0}' AND c.fecha BETWEEN STR_TO_DATE('{1}', '%m/%d/%Y') AND STR_TO_DATE('{2}', '%m/%d/%Y') ORDER BY fecha ASC;".format(rut,fechaInicial,fechaFinal)

    
    results = None
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
    except:
        print("error en la base")
        db.rollback()
        Desconectar(db)
        return
    
    return results

def ObtenerMayor(rut,codigoCuenta,fechaInicial,fechaFinal):
    db = Conectar()
    
    cursor = db.cursor()
    sql = None
    if codigoCuenta == "todos":
        sql = "SELECT c.fecha,c.numMes,c.tipo,c.glosa,t.debe,t.haber,cu.codigo,cu.nombre FROM comprobante c inner join transaccion t inner join cuenta cu on c.codigoCom = t.codigoCom AND t.codigo = cu.codigo AND c.rut = t.rut AND cu.rut = t.rut WHERE c.rut = '{0}' AND c.fecha BETWEEN STR_TO_DATE('{1}', '%m/%d/%Y') AND STR_TO_DATE('{2}', '%m/%d/%Y') ORDER BY fecha ASC;".format(rut,fechaInicial,fechaFinal)
    else:
        sql = "SELECT c.fecha,c.numMes,c.tipo,c.glosa,t.debe,t.haber,cu.codigo,cu.nombre FROM comprobante c inner join transaccion t inner join cuenta cu on c.codigoCom = t.codigoCom AND t.codigo = cu.codigo AND c.rut = t.rut AND cu.rut = t.rut WHERE c.rut = '{0}' AND cu.codigo = '{1}' AND c.fecha BETWEEN STR_TO_DATE('{2}', '%m/%d/%Y') AND STR_TO_DATE('{3}', '%m/%d/%Y') ORDER BY fecha ASC;".format(rut,codigoCuenta,fechaInicial,fechaFinal)
    
    results = None
    
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
    except:
        print("error en la base")
        db.rollback()
        Desconectar(db)
        return
    
    return results
    
def VerDiario(rut,fechaInicial,fechaFinal):
    
    db = Conectar()
    cursor = db.cursor()
    sql = "SELECT fecha,codigoCom,tipo,glosa FROM comprobante WHERE rut = '{0}' AND fecha BETWEEN STR_TO_DATE('{1}', '%m/%d/%Y') AND STR_TO_DATE('{2}', '%m/%d/%Y') ORDER BY fecha ASC;".format(rut,fechaInicial,fechaFinal)
    results = None
    i = 1
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
    except:
        print("error en la base")
        db.rollback()
        Desconectar(db)
        return
    
    for row in results:
        sql = "SELECT debe,haber,codigo FROM transaccion WHERE rut = '{0}' AND codigoCom = '{1}';".format(rut,row[1])
        results2 = None
        try:
            cursor.execute(sql)
            results2 = cursor.fetchall()
        except:
            print("error en la base")
            db.rollback()
            Desconectar(db)
            return
        
        
        print("comprobante: '{0}'".format(i))
        print("fecha: '{0}'".format(row[0]))
        print("tipo: '{0}'".format(row[2]))
        print("glosa: '{0}'").format(row[3])
        
        
        j = 1
        for row2 in results2:
            
            sql = "SELECT nombre,tipo FROM cuenta WHERE rut = '{0}' AND codigo = '{1}';".format(rut,row2[2])
            results3 = None
            
            try:
                cursor.execute(sql)
                results3 = cursor.fetchall()
            except:
                print("error en la base")
                db.rollback()
                Desconectar(db)
                return
            
            print("transaccion '{0}'".format(j))
            print("debe: '{0}'".format(row2[0]))
            print("haber: '{0}'".format(row2[1]))
            print("nombre cuenta: '{0}'".format(results3[0][0]))
            print("tipo cuenta: '{0}'".format(results3[0][1]))
            print("#############################")
            
            j += 1
            
        i += 1

def ObtenerBalance(rut,fechaInicial,fechaFinal):
    db = Conectar()
    
    cursor = db.cursor()
    
    sql = "SELECT SUM(debe), SUM(haber), nombre, cu.codigo, cu.tipo FROM comprobante c inner join transaccion t inner join cuenta cu on c.codigoCom = t.codigoCom AND t.codigo = cu.codigo AND c.rut = t.rut AND cu.rut = t.rut WHERE c.rut = '{0}' AND c.fecha BETWEEN STR_TO_DATE('{1}', '%m/%d/%Y') AND STR_TO_DATE('{2}', '%m/%d/%Y') GROUP BY cu.nombre, cu.codigo, cu.tipo;".format(rut,fechaInicial,fechaFinal)
    results = None
    
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
    except:
        print("error en la base")
        db.rollback()
        Desconectar(db)
        return
    
    return results        


def VerBalance(rut,fechaInicial,fechaFinal):
    print("balance")
    
    db = Conectar()
    
    cursor = db.cursor()
    
    sql = "SELECT codigo, nombre, tipo FROM cuenta WHERE rut = '{0}' ORDER BY codigo ASC;".format(rut)
    
    results = None
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
    except:
        print("error en la base")
        db.rollback()
        Desconectar(db)
        return
    
    
    
    for row in results:
        
        sql = "SELECT SUM(debe), SUM(haber), codigoCom FROM transaccion AS T WHERE codigo = '{0}' AND rut = '{1}' AND (SELECT fecha FROM comprobante WHERE codigoCom = T.codigoCom) BETWEEN STR_TO_DATE('{2}', '%m/%d/%Y') AND STR_TO_DATE('{3}', '%d/%m/%Y')".format(row[0],rut,fechaInicial,fechaFinal)
        
        results2 = None
        try:
            cursor.execute(sql)
            results2 = cursor.fetchall()
        except:
            print("error en la base")
            db.rollback()
            Desconectar(db)
            return
        
        print("codigo: '{0}'".format(row[0]))
        print("nombre: '{0}'".format(row[1]))
        print("debe: '{0}'".format(results2[0][0]))
        print("haber: '{0}'".format(results2[0][1]))
        
        if results2[0][0] == None or results2[0][1] == None:
            return
        
        deudor = results2[0][0] - results2[0][1]
        acreedor = results2[0][1] - results2[0][0]
        
        if deudor < 0:
            deudor = 0
        
        if acreedor < 0:
            acreedor = 0
        
        activo = 0
        pasivo = 0
        perdida = 0
        ganancia = 0
        
        if row[2] == "activo" or row[2] == "pasivo" or row[2] == "patrimonio":
            activo = deudor
            pasivo = acreedor
        else:
            perdida = deudor
            ganancia = acreedor
        
        print("saldos: deudor = '{0}'; acreedor = '{1}'".format(deudor, acreedor))
        print("inventario: activo = '{0}'; pasivo = '{1}'".format(activo, pasivo))
        print("resultado: perdida = '{0}'; ganancia = '{1}'".format(perdida, ganancia))
        
        
    
            
        
    ['Are', 'there', 'any', 'eritrean', 'restaurants', 'in', 'town', '?']
    
    
def ObtenerDatosParaExcel(rut):
    db = Conectar()
    
    cursor = db.cursor()
    sql = "SELECT nombre, direccion, giro, dueño, rutProp FROM empresa WHERE rut = '{0}'".format(rut)
     
    results = None
    try:
        cursor.execute(sql)
        results = cursor.fetchall()
    except:
        print("error en la base")
        db.rollback()
        Desconectar(db)
        return
    
    return results[0][0], results[0][1], results[0][2], results[0][3], results[0][4]
    
"""def ObtenerDatosParaExcelDiario(rut):
    db = Conectar()
    
    cursor = db.cursor()"""