using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;

namespace SeleniumScrapper
{
    class Program
    {
        static void Main(string[] args)
        {
            String[] instrucciones = new string[6] {
                "Inicie sesión con sus credenciales, luego dirigase a Alumnos>Ficha alumno y espere a que cargue.",
                "Seleccione los parámetros para 'Tipos de enseñanza', 'Grados' y 'Cursos'.",
                "Haga click en 'Buscar'",
                "Inicie sesión con sus credenciales y espere a que cargue la página.",
                "Se cargará la página 'nombrerutyfirma', por favor espere y/o siga los pasos hasta llegar al buscador.",
                "Después presione una tecla para continuar."
            };
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.WriteLine("IMPORTANTE: ");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine("La función de este programa es específica (No es dinámico). ");
            Console.WriteLine("Cada vez que termine una instrucción, debe presionar una tecla en esta ventana.");
            Console.Write("Es posible que en unos meses o años más ");
            Console.WriteLine("este ya no cumpla su objetivo debido a cambios en los sitios web destinados.");
            Console.WriteLine("");
            Console.WriteLine("Para empezar presione una tecla.");
            Console.ReadKey();
            Console.Clear();

            char opcion;

            ChromeDriver driver = new ChromeDriver();
            IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
            bool inicioSesionNapsis = false;
            bool inicioSesionSige = false;

            IJavaScriptExecutor js = (IJavaScriptExecutor)driver;

            do
            {
                Console.Clear();
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Menú");
                Console.ForegroundColor = ConsoleColor.White;
                Console.WriteLine("1: Copiar información de alumnos desde Napsis");
                Console.WriteLine("2: Traspasar información de alumnos hacia SIGE");
                Console.WriteLine("3: Listar información respaldada de alumnos");
                Console.WriteLine("4: Listar información respaldada de un alumno específico");
                Console.WriteLine("5: Salir");

                opcion = Console.ReadKey().KeyChar;
                Console.WriteLine("");

                var archivosTxt = Directory.EnumerateFiles("C:/Alumnos/", "*.txt");
                var archivosTxtFamiliares = Directory.EnumerateFiles("C:/Alumnos/", "F*");

                switch (opcion)
                {
                    case '1':
                        if (!inicioSesionNapsis)
                        {
                            driver.Url = "https://napsis.com/ingresar/";
                            CWInstrucciones();
                            Console.WriteLine(instrucciones[0]);
                            Console.WriteLine(instrucciones[5]);
                            Console.ReadKey();
                            inicioSesionNapsis = true;
                        }
                        
                        driver.Url = "http://ncore.napsis.cl/fichaalumno/ficha-alumno";
                        CWInstrucciones();
                        Console.WriteLine(instrucciones[1]);
                        Console.WriteLine(instrucciones[2]);
                        Console.WriteLine(instrucciones[5]);
                        Console.WriteLine("");
                        Console.WriteLine("Presione una tecla para continuar.");
                        Console.ReadKey();

                        int total = Convert.ToInt32(driver.ExecuteScript("return document.querySelector('#tabla_listado_alumnos > tbody').rows.length"));

                        Console.WriteLine("Alumnos en esta lista: " + total);
                        Console.WriteLine("");

                        string rut, nombre, apaterno, amaterno, fechanac, direc, comuna, sexo, celular, nacionalidad;
                        string rutF, nombreF, apaternoF, amaternoF, fechanacF, direcF, comunaF, sexoF, telefonoF, celularF, nacionalidadF, parentescoF;
                        int cantidadFamiliares;
                        string path;


                        for (int alumnos = 1; alumnos <= total; alumnos++)
                        {
                            wait.Until(ExpectedConditions.ElementExists(By.ClassName("odd")));
                            driver.ExecuteScript("document.getElementsByClassName('btn btn-small btn-very-tiny ')[" + (alumnos - 1) + "].click();");

                            wait.Until(ExpectedConditions.ElementExists(By.Id("campo_1_8576")));
                            rut = Convert.ToString(driver.FindElementById("campo_1_8576").GetAttribute("Value"));
                            nombre = Convert.ToString(driver.FindElementById("campo_2_8576").GetAttribute("Value"));
                            apaterno = Convert.ToString(driver.FindElementById("campo_4_8576").GetAttribute("Value"));
                            amaterno = Convert.ToString(driver.FindElementById("campo_3_8576").GetAttribute("Value"));
                            fechanac = Convert.ToString(driver.FindElementById("campo_6_8576").GetAttribute("Value"));
                            direc = Convert.ToString(driver.FindElementById("campo_8_8576").GetAttribute("Value"));
                            comuna = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[4].text;"));
                            sexo = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[5].text;"));
                            celular = Convert.ToString(driver.FindElementById("campo_9_8576").GetAttribute("Value"));
                            nacionalidad = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[6].text;"));


                            driver.ExecuteScript("cambiarTab(3);");
                            wait.Until(ExpectedConditions.ElementExists(By.Id("accordion")));
                            cantidadFamiliares = Convert.ToInt32(driver.ExecuteScript("return document.getElementById('accordion').childElementCount")) / 2;

                            if (cantidadFamiliares > 0)
                            {
                                for (int familiar = 1; familiar <= cantidadFamiliares; familiar++)
                                {
                                    switch (familiar)
                                    {
                                        case 1:
                                            rutF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[28].children[0].value;"));
                                            nombreF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[30].children[0].value;"));
                                            apaternoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[31].children[0].value;"));
                                            amaternoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[32].children[0].value;"));
                                            fechanacF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[42].children[0].value;"));
                                            direcF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[38].children[0].children[0].value;"));
                                            comunaF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[16].children[0].textContent;"));
                                            sexoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[13].children[0].textContent;"));
                                            celularF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[34].children[0].value;"));
                                            telefonoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[33].children[0].value;"));
                                            nacionalidadF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[17].children[0].textContent;"));
                                            parentescoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[15].children[0].textContent;"));
                                            break;
                                        case 2:
                                            rutF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[52].children[0].value;"));
                                            nombreF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[54].children[0].value;"));
                                            apaternoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[55].children[0].value;"));
                                            amaternoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[56].children[0].value;"));
                                            fechanacF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[66].children[0].value;"));
                                            direcF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[62].children[0].children[0].value;"));
                                            comunaF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[24].children[0].textContent;"));
                                            sexoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[21].children[0].textContent;"));
                                            celularF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[58].children[0].value;"));
                                            telefonoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[57].children[0].value;"));
                                            nacionalidadF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[25].children[0].textContent;"));
                                            parentescoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[23].children[0].textContent;"));
                                            break;
                                        case 3:
                                            rutF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[76].children[0].value;"));
                                            nombreF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[78].children[0].value;"));
                                            apaternoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[79].children[0].value;"));
                                            amaternoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[80].children[0].value;"));
                                            fechanacF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[90].children[0].value;"));
                                            direcF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[86].children[0].children[0].value;"));
                                            comunaF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[32].children[0].textContent;"));
                                            sexoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[29].children[0].textContent;"));
                                            celularF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[82].children[0].value;"));
                                            telefonoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('controls')[81].children[0].value;"));
                                            nacionalidadF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[33].children[0].textContent;"));
                                            parentescoF = Convert.ToString(driver.ExecuteScript("return document.getElementsByClassName('select2-choice')[31].children[0].textContent;"));
                                            break;
                                        default:
                                            rutF = "Llenar a mano";
                                            nombreF = "Llenar a mano";
                                            apaternoF = "Llenar a mano";
                                            amaternoF = "Llenar a mano";
                                            fechanacF = "Llenar a mano";
                                            direcF = "Llenar a mano";
                                            comunaF = "Llenar a mano";
                                            sexoF = "Llenar a mano";
                                            celularF = "Llenar a mano";
                                            telefonoF = "Llenar a mano";
                                            nacionalidadF = "Llenar a mano";
                                            parentescoF = "Llenar a mano";
                                            break;
                                    }

                                    path = @"c:\Alumnos\F" + familiar + " " + rut + ".txt";
                                    if (!File.Exists(path))
                                    {
                                        using (StreamWriter sw = File.CreateText(path))
                                        {
                                            sw.WriteLine(rutF);
                                            sw.WriteLine(nombreF.Trim());
                                            sw.WriteLine(apaternoF);
                                            sw.WriteLine(amaternoF);
                                            sw.WriteLine(fechanacF);
                                            sw.WriteLine(direcF);
                                            sw.WriteLine(comunaF);
                                            sw.WriteLine(sexoF);
                                            sw.WriteLine(celularF);
                                            sw.WriteLine(telefonoF);
                                            sw.WriteLine(nacionalidadF);
                                            sw.WriteLine(parentescoF);
                                        }
                                        Console.WriteLine("[" + familiar + "/" + cantidadFamiliares + "] " + "Información del familiar " + nombreF + " guardada.");
                                    }

                                }
                            }

                            path = @"c:\Alumnos\" + rut + ".txt";
                            if (!File.Exists(path))
                            {
                                using (StreamWriter sw = File.CreateText(path))
                                {
                                    sw.WriteLine(rut);
                                    sw.WriteLine(nombre.Trim());
                                    sw.WriteLine(apaterno);
                                    sw.WriteLine(amaterno);
                                    sw.WriteLine(fechanac);
                                    sw.WriteLine(direc);
                                    sw.WriteLine(comuna.Trim());
                                    sw.WriteLine(sexo.Trim());
                                    sw.WriteLine(celular);
                                    sw.WriteLine(nacionalidad.Trim());
                                    sw.WriteLine(cantidadFamiliares);
                                }
                                Console.WriteLine("["+ alumnos + "/" + total + "] " +  "Información de " + nombre + " guardada.");
                            }

                            wait.Until(ExpectedConditions.ElementExists(By.Id("campo_1_8576")));
                            Thread.Sleep(1000);
                            driver.ExecuteScript("guardarBreadcrumbOut();");
                        }
                        break;
                    case '2':
                        String[] texto;
                        List<String> familiaresSinRC = new List<string>();
                        if (archivosTxt.Count() > 0) //if existe más de un archivo en C:/Alumnos
                        {
                            Console.WriteLine("");
                            Console.WriteLine("Este proceso ingresará toda la información almacenada en C:/Alumnos utilizando su cuenta.");
                            Console.WriteLine("Si está seguro de continuar, presione una tecla. Sino, cierre esta ventana.");
                            Console.ReadKey();

                            CWInstrucciones();
                            Console.WriteLine(instrucciones[4]);
                            Console.WriteLine(instrucciones[5]);
                            driver.Url = "https://nombrerutyfirma.com";
                            Console.ReadKey();

                            if (!inicioSesionSige)
                            {
                                driver.Url = "https://sige.mineduc.cl/Sige/Login";
                                CWInstrucciones();
                                Console.WriteLine(instrucciones[3]);
                                Console.WriteLine(instrucciones[5]);
                                Console.ReadKey();
                                inicioSesionSige = true;
                            }

                            wait.Until(ExpectedConditions.ElementExists(By.Id("ui-dialog-title-solidario")));
                            driver.ExecuteScript("document.getElementsByTagName('a')[20].click();");
                            driver.ExecuteScript("document.getElementsByTagName('a')[5].click();");

                            wait.Until(ExpectedConditions.ElementExists(By.ClassName("nivel3")));
                            int cantidadCursos = Convert.ToInt32(driver.ExecuteScript("return document.getElementsByTagName('td').length"));
                            

                            js.ExecuteScript("window.open();");
                            IList<string> tabs = new List<string>(driver.WindowHandles);
                            driver.SwitchTo().Window(tabs[0]);

                            for (int i = 0; i <= cantidadCursos-1; i++)
                            {
                                driver.ExecuteScript("document.getElementsByTagName('a')[12].click();");
                                wait.Until(ExpectedConditions.ElementExists(By.ClassName("tblCursosTDLigth")));
                                Thread.Sleep(100);

                                int cantidadAlumnos = Convert.ToInt32(Convert.ToString(driver.ExecuteScript("return document.getElementsByTagName('p')[" + i + 2 + "].children[0].text;")).Substring(12, 2));
                                driver.ExecuteScript("document.getElementsByTagName('p')["+i+2+"].children[0].click();");
                                
                                foreach (string archivo in archivosTxt)
                                {
                                    for (int a = 0; a < cantidadAlumnos-1; a++)
                                    {
                                        wait.Until(ExpectedConditions.ElementExists(By.ClassName("nivel"+cantidadAlumnos)));
                                        Thread.Sleep(300);
                                        int idElementoRutAlumno = 18 + (2 * a);
                                        string rutAlumnoActual = Convert.ToString(driver.ExecuteScript("return document.getElementsByTagName('a')["+idElementoRutAlumno+"].text;")).Replace(".", "");
                                        if (archivo.Contains(rutAlumnoActual))
                                        {
                                            string[] lineasTexto = File.ReadAllLines(archivo);
                                            driver.ExecuteScript("document.getElementsByTagName('a')[" + idElementoRutAlumno + "].click();");

                                            //Información del alumno
                                            wait.Until(ExpectedConditions.ElementExists(By.Id("chkSolicitaVal")));
                                            driver.ExecuteScript("document.getElementById('chkSolicitaVal').click();");
                                            Thread.Sleep(200);
                                            IAlert alertOk = driver.SwitchTo().Alert();
                                            alertOk.Accept();

                                            driver.ExecuteScript("document.getElementsByName('cmbRegion')[0].children[16].selected = true");
                                            driver.ExecuteScript("cambiaComunas(13);");
                                            for (int c = 0; c <= 52; c++)
                                            {
                                                string opcionComunaActual = Convert.ToString(driver.ExecuteScript("return document.getElementsByName('cmbComuna')[0].children[" + c + "].text.toLowerCase();"));
                                                if (lineasTexto[6].ToLower() == opcionComunaActual)
                                                {
                                                    driver.ExecuteScript("document.getElementsByName('cmbComuna')[0].children[" + c + "].selected = true;");
                                                }
                                            }
                                            driver.ExecuteScript("document.getElementsByName('txtDir')[0].value = '" + lineasTexto[5] + "';");
                                            driver.ExecuteScript("document.getElementsByName('txtFono')[0].value = '" + lineasTexto[8] + "';");

                                            driver.ExecuteScript("document.getElementsByClassName('boton_actualizar_mediano')[0].click();");
                                            wait.Until(ExpectedConditions.ElementExists(By.Id("popup_ok")));
                                            driver.ExecuteScript("document.getElementById('popup_ok').click()");
                                            driver.Navigate().Back();

                                            bool tabFamiliarCargado = false;

                                            string[] archivoSplit = archivo.Split(".");
                                            foreach(string archivoFamiliar in archivosTxtFamiliares)
                                            {
                                                if (archivoFamiliar.Contains(lineasTexto[0]))
                                                {
                                                    //Información de familiares
                                                    if (!tabFamiliarCargado)
                                                    {
                                                        wait.Until(ExpectedConditions.ElementExists(By.ClassName("nivel1")));
                                                        driver.ExecuteScript("document.getElementsByTagName('a')[16].click();");
                                                        tabFamiliarCargado = true;
                                                    }
                                                    
                                                    wait.Until(ExpectedConditions.ElementExists(By.Id("txtRunFamiliar")));

                                                    string[] lineasTextoFamiliar = File.ReadAllLines(archivoFamiliar);
                                                    
                                                    driver.SwitchTo().Window(tabs[1]);
                                                    driver.Url = "https://www.nombrerutyfirma.com/";

                                                    wait.Until(ExpectedConditions.ElementExists(By.ClassName("text-center")));
                                                    driver.ExecuteScript("document.getElementsByTagName('a')[2].click();");
                                                    Thread.Sleep(100);

                                                    string[] rutFamiliarDividido = lineasTextoFamiliar[0].Split("-");
                                                    rutFamiliarDividido[0] = Reverse(rutFamiliarDividido[0]);
                                                    rutFamiliarDividido[0] = rutFamiliarDividido[0].Substring(0, 3) + "." + rutFamiliarDividido[0].Substring(3, 3) + "." + rutFamiliarDividido[0].Substring(6, rutFamiliarDividido[0].Length-6);
                                                    rutFamiliarDividido[0] = Reverse(rutFamiliarDividido[0]);
                                                    if (rutFamiliarDividido[1].ToLower() == "k")
                                                    {
                                                        driver.ExecuteScript("document.getElementsByName('term')[1].value = '" + rutFamiliarDividido[0] + "-0';");
                                                    }
                                                    else
                                                    {
                                                        driver.ExecuteScript("document.getElementsByName('term')[1].value = '" + rutFamiliarDividido[0] + "-" + rutFamiliarDividido[1] + "';");
                                                    }
                                                    
                                                    Thread.Sleep(100);
                                                    driver.ExecuteScript("document.getElementsByClassName('btn btn-primary')[1].click();");

                                                    wait.Until(ExpectedConditions.ElementExists(By.ClassName("table-hover")));
                                                    int numeroFilas = Convert.ToInt32(driver.ExecuteScript("return document.getElementsByClassName('table table-hover')[0].rows.length"));
                                                    if (numeroFilas == 2)
                                                    {
                                                        string[] arrayNombreFamiliar = Convert.ToString(driver.ExecuteScript("return document.getElementsByTagName('td')[0].textContent;")).Split(" ");
                                                        string direccionFamiliarRutificador = Convert.ToString(driver.ExecuteScript("return document.getElementsByTagName('td')[3].textContent;"));
                                                        driver.SwitchTo().Window(tabs[0]);

                                                        for (int p = 0; p <= 5; p++)
                                                        {
                                                            string opcionParentescoActual = Convert.ToString(driver.ExecuteScript("return document.getElementsByName('cmbTipoFamiliar')[0].children[" + p + "].textContent;"));
                                                            if (opcionParentescoActual == lineasTextoFamiliar[11])
                                                            {
                                                                driver.ExecuteScript("document.getElementsByName('cmbTipoFamiliar')[0].children[" + p + "].selected = true;");
                                                            }
                                                        }
                                                        string[] rutDividido = lineasTextoFamiliar[0].Split("-");
                                                        driver.ExecuteScript("document.getElementsByName('txtRunFamiliar')[0].value = '" + rutDividido[0] + "';");
                                                        driver.ExecuteScript("document.getElementsByName('txtDVRunFamiliar')[0].value = '" + rutDividido[1] + "';");

                                                        driver.ExecuteScript("document.getElementsByName('txtPerNombre')[0].value = '" + arrayNombreFamiliar[2] + " " + arrayNombreFamiliar[3] + "';");
                                                        driver.ExecuteScript("document.getElementsByName('txtPerPaterno')[0].value = '" + arrayNombreFamiliar[0] + "';");
                                                        driver.ExecuteScript("document.getElementsByName('txtPerMaterno')[0].value = '" + arrayNombreFamiliar[1] + "';");

                                                        for (int pp = 0; pp <= 28; pp++)
                                                        {
                                                            wait.Until(ExpectedConditions.ElementExists(By.Id("cmbPerPaisOrigen")));
                                                            string opcionPaisProcedenciaActual = Convert.ToString(driver.ExecuteScript("return document.getElementsByName('cmbPerPaisOrigen')[0].children[" + pp + "].text;")).ToLower();
                                                            if (lineasTextoFamiliar[10].Length > 0)
                                                            {
                                                                if (opcionPaisProcedenciaActual.Contains(lineasTextoFamiliar[10].Substring(0, 4)))
                                                                {
                                                                    driver.ExecuteScript("document.getElementsByName('cmbPerPaisOrigen')[0].children[" + pp + "].selected = true;");
                                                                }
                                                            }
                                                        }
                                                        for (int s = 0; s <= 2; s++)
                                                        {
                                                            string opcionSexoActual = Convert.ToString(driver.ExecuteScript("return document.getElementsByName('cmbSexoFamiliar')[0].children[" + s + "].text.toLowerCase(); "));
                                                            if (opcionSexoActual == lineasTextoFamiliar[7].ToLower())
                                                            {
                                                                driver.ExecuteScript("document.getElementsByName('cmbSexoFamiliar')[0].children[" + s + "].selected = true;");
                                                            }
                                                        }
                                                        driver.ExecuteScript("document.getElementsByName('cmbRegion')[0].children[16].selected = true");
                                                        driver.ExecuteScript("cambiaComunas(13);");
                                                        for (int c = 0; c <= 52; c++)
                                                        {
                                                            string opcionComunaActual = Convert.ToString(driver.ExecuteScript("return document.getElementsByName('cmbComuna')[0].children[" + c + "].text.toLowerCase();"));
                                                            if (lineasTextoFamiliar[6].ToLower() == opcionComunaActual)
                                                            {
                                                                driver.ExecuteScript("document.getElementsByName('cmbComuna')[0].children[" + c + "].selected = true;");
                                                            }
                                                        }

                                                        var numeroDireccion = new StringBuilder();

                                                        bool segundoIf = false;
                                                        bool primerIf = true;

                                                        foreach (var charDirecc in direccionFamiliarRutificador)
                                                        {
                                                            if (char.IsDigit(charDirecc) && primerIf == true)
                                                            {
                                                                numeroDireccion.Append(charDirecc);
                                                                segundoIf = true;
                                                            }
                                                            else
                                                            {
                                                                if (segundoIf)
                                                                {
                                                                    if (!char.IsDigit(charDirecc))
                                                                    {
                                                                        primerIf = false;
                                                                        segundoIf = false;
                                                                    }
                                                                }
                                                            }
                                                        }

                                                        driver.ExecuteScript("document.getElementsByName('txtDir')[0].value = '" + direccionFamiliarRutificador.Replace(numeroDireccion.ToString(), "") + "';");
                                                        driver.ExecuteScript("document.getElementsByName('txtDirNumero')[0].value = '" + numeroDireccion + "';");
                                                        driver.ExecuteScript("document.getElementsByName('txtCelularFamiliar')[0].value = '" + lineasTextoFamiliar[8] + "';");
                                                        //Implementar si es necesario agregar nivel educacional y ocupación actual
                                                        /*for (int e = 0; e <= 20; e++)
                                                        {
                                                            string opcionEducacionActual = Convert.ToString(driver.ExecuteScript("return document.getElementsByName('cmbUltNivEducFamiliar')[0].children["+e+"].text"));
                                                            if (lineasTextoFamiliar[lineaEducacionActual] == opcionEducacionActual)
                                                            {
                                                                driver.ExecuteScript("document.getElementsByName('cmbUltNivEducFamiliar')[0].children[" + e + "].selected = true;");
                                                            }
                                                        }
                                                        for (int o = 0; o <= 5; o++)
                                                        {
                                                            string opcionOcupacionActual = Convert.ToString(driver.ExecuteScript("return document.getElementsByName('cmbSitLaborFamiliar')[0].children["+o+"].text"));
                                                            if (lineasTextoFamiliar[lineaOcupacionActual] == opcionOcupacionActual)
                                                            {
                                                                driver.ExecuteScript("document.getElementsByName('cmbSitLaborFamiliar')[0].children["+o+"].selected = true;");
                                                            }
                                                        }*/
                                                        driver.ExecuteScript("document.getElementsByName('cmbUltNivEducFamiliar')[0].children[20].selected = true;");
                                                        driver.ExecuteScript("document.getElementsByName('cmbSitLaborFamiliar')[0].children[5].selected = true;");
                                                        driver.ExecuteScript("document.getElementById('rdoTrabajaFamiliar2').click()");
                                                        
                                                        driver.ExecuteScript("document.getElementsByClassName('boton_ingresar_mediano')[0].click();");
                                                        driver.ExecuteScript("document.getElementById('popup_ok').click();");
                                                    }
                                                    else
                                                    {
                                                        driver.SwitchTo().Window(tabs[0]);
                                                        familiaresSinRC.Add(archivoFamiliar);
                                                        continue;
                                                    }
                                                }
                                            }
                                            wait.Until(ExpectedConditions.ElementExists(By.ClassName("boton_anterior_mediano")));
                                            driver.ExecuteScript("document.getElementsByClassName('boton_anterior_mediano')[0].click();");
                                        }
                                    }
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("No existen archivos para copiar a SIGE");
                            Console.WriteLine("Oprima una tecla para volver al menú.");
                            Console.ReadKey();
                        }
                        Console.WriteLine("");
                        Console.WriteLine("Formato: F<numeroDeFamiliar> <rutAlumno>.txt");
                        Console.WriteLine("Los siguientes familiares no pudieron ser confirmados a través de Rutificador:");
                        foreach (string familiar in familiaresSinRC)
                        {
                            Console.ForegroundColor = ConsoleColor.Red;
                            Console.WriteLine(familiar);
                        }
                        Console.ForegroundColor = ConsoleColor.White;
                        Console.WriteLine("");
                        Console.WriteLine("Presione una tecla para continuar.");
                        Console.ReadKey();
                        break;
                    case '3':
                        Console.WriteLine("");
                        foreach (string archivo in archivosTxt)
                        {
                            texto = File.ReadAllLines(archivo);
                            int contador = 0;

                            if (archivo.Contains("F"))
                            {
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("--- Familiar " + texto[1] + " (" + archivo + ") ---");
                                Console.ForegroundColor = ConsoleColor.White;
                                foreach (string linea in texto)
                                {
                                    switch (contador)
                                    {
                                        case 0:
                                            Console.WriteLine("Rut:           " + linea);
                                            break;
                                        case 1:
                                            Console.WriteLine("Nombres:       " + linea);
                                            break;
                                        case 2:
                                            Console.WriteLine("A. Paterno:    " + linea);
                                            break;
                                        case 3:
                                            Console.WriteLine("A. Materno:    " + linea);
                                            break;
                                        case 4:
                                            Console.WriteLine("F. Nacimiento: " + linea);
                                            break;
                                        case 5:
                                            Console.WriteLine("Dirección:     " + linea);
                                            break;
                                        case 6:
                                            Console.WriteLine("Comuna:        " + linea);
                                            break;
                                        case 7:
                                            Console.WriteLine("Sexo:          " + linea);
                                            break;
                                        case 8:
                                            Console.WriteLine("Celular:       " + linea);
                                            break;
                                        case 9:
                                            Console.WriteLine("Telefono:      " + linea);
                                            break;
                                        case 10:
                                            Console.WriteLine("Nacionalidad:  " + linea);
                                            break;
                                        case 11:
                                            Console.WriteLine("Parentesco:    " + linea);
                                            break;
                                    }
                                    contador++;
                                }
                            }
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("--- Alumno " + texto[1] + " ---");
                                Console.ForegroundColor = ConsoleColor.White;
                                foreach (string linea in texto)
                                {
                                    switch (contador)
                                    {
                                        case 0:
                                            Console.WriteLine("Rut:           " + linea);
                                            break;
                                        case 1:
                                            Console.WriteLine("Nombres:       " + linea);
                                            break;
                                        case 2:
                                            Console.WriteLine("A. Paterno:    " + linea);
                                            break;
                                        case 3:
                                            Console.WriteLine("A. Materno:    " + linea);
                                            break;
                                        case 4:
                                            Console.WriteLine("F. Nacimiento: " + linea);
                                            break;
                                        case 5:
                                            Console.WriteLine("Dirección:     " + linea);
                                            break;
                                        case 6:
                                            Console.WriteLine("Comuna:        " + linea);
                                            break;
                                        case 7:
                                            Console.WriteLine("Sexo:          " + linea);
                                            break;
                                        case 8:
                                            Console.WriteLine("Celular:       " + linea);
                                            break;
                                        case 9:
                                            Console.WriteLine("Nacionalidad:  " + linea);
                                            break;
                                        case 10:
                                            Console.WriteLine("C. Familiares: " + linea);
                                            break;
                                    }
                                    contador++;
                                }
                            }
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("---------------------------");
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine("");
                        }
                        Console.WriteLine("Oprima una tecla para volver al menú.");
                        Console.ReadKey();
                        break;
                    case '4':
                        Console.WriteLine("Digite el rut del alumno sin puntos");
                        Console.WriteLine("Ejemplo: 12345678-9");
                        Console.WriteLine("Rut: ");
                        string rutBusqueda = Console.ReadLine();
                        Console.WriteLine("");

                        int cantidadResultados = 0;

                        if (archivosTxt.Count() > 0)
                        {
                            foreach (string archivo in archivosTxt)
                            {
                                texto = File.ReadAllLines(archivo);
                                
                                if (texto[0].Contains(rutBusqueda))
                                {
                                    cantidadResultados++;
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("--- Alumno " + texto[1] + " " + texto[2] + " " + texto[3] + " ---");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.WriteLine("Rut:           " + texto[0]);
                                    Console.WriteLine("Nombres:       " + texto[1]);
                                    Console.WriteLine("A. Paterno:    " + texto[2]);
                                    Console.WriteLine("A. Materno:    " + texto[3]);
                                    Console.WriteLine("F. Nacimiento: " + texto[4]);
                                    Console.WriteLine("Dirección:     " + texto[5]);
                                    Console.WriteLine("Comuna:        " + texto[6]);
                                    Console.WriteLine("Sexo:          " + texto[7]);
                                    Console.WriteLine("Celular:       " + texto[8]);
                                    Console.WriteLine("Nacionalidad:  " + texto[9]);
                                    Console.WriteLine("C. Familiares: " + texto[10]);
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("---------------------------");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.WriteLine("");
                                }

                                if (archivo.Contains(rutBusqueda) && archivo.Contains("F"))
                                {
                                    cantidadResultados++;
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("--- Familiar " + texto[1] + " (" + archivo + ") ---");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.WriteLine("Rut:           " + texto[0]);
                                    Console.WriteLine("Nombres:       " + texto[1]);
                                    Console.WriteLine("A. Paterno:    " + texto[2]);
                                    Console.WriteLine("A. Materno:    " + texto[3]);
                                    Console.WriteLine("F. Nacimiento: " + texto[4]);
                                    Console.WriteLine("Dirección:     " + texto[5]);
                                    Console.WriteLine("Comuna:        " + texto[6]);
                                    Console.WriteLine("Sexo:          " + texto[7]);
                                    Console.WriteLine("Celular:       " + texto[8]);
                                    Console.WriteLine("Telefono:      " + texto[9]);
                                    Console.WriteLine("Nacionalidad:  " + texto[10]);
                                    Console.WriteLine("Parentesco:    " + texto[11]);
                                    Console.ForegroundColor = ConsoleColor.Green;
                                    Console.WriteLine("---------------------------");
                                    Console.ForegroundColor = ConsoleColor.White;
                                    Console.WriteLine("");
                                }
                            }
                            if (cantidadResultados == 0) { Console.WriteLine("No se encontraron resultados."); }
                        }
                        Console.WriteLine("Oprima una tecla para volver al menú.");
                        Console.ReadKey();
                        break;
                    case '5':
                        driver.Dispose();
                        Environment.Exit(0);
                        break;
                }
                
            } while (opcion != '5');
        }

        private static void CWInstrucciones()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Instrucciones:");
            Console.ForegroundColor = ConsoleColor.Yellow;
        }
        public static string Reverse(string s)
        {
            char[] charArray = s.ToCharArray();
            Array.Reverse(charArray);
            return new string(charArray);
        }
    }
}