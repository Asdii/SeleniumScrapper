using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using System;
using System.IO;
using System.Linq;
using System.Threading;

namespace SeleniumScrapper
{
    class Program
    {
        static void Main(string[] args)
        {
            String[] instrucciones = new string[4] {
                "Inicie sesión con sus credenciales, luego dirigase a Alumnos>Ficha alumno y espere a que cargue.",
                "Seleccione los parámetros para 'Tipos de enseñanza', 'Grados' y 'Cursos'.",
                "Haga click en 'Buscar'",
                "Inicie sesión con sus credenciales, espere a que cargue la página y presione una tecla para continuar."
            };

            Thread.Sleep(500);
            Console.ForegroundColor = ConsoleColor.Yellow;
            Thread.Sleep(500);
            Console.WriteLine("IMPORTANTE: ");
            Console.ForegroundColor = ConsoleColor.White;
            Thread.Sleep(500);
            Console.WriteLine("La función de este programa es específica (No es dinámico). ");
            Console.WriteLine("Cada vez que termine una instrucción, debe presionar una tecla en esta ventana.");
            Console.Write("Es posible que en unos meses o años más ");
            Console.WriteLine("este ya no cumpla su objetivo debido a cambios en los sitios web destinados.");
            Console.WriteLine("");
            Console.Write("Toda instrucción debe ser seguida, los pasos en ");
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("amarillo");
            Console.ForegroundColor = ConsoleColor.White;
            Console.WriteLine(" son necesarios para continuar con el funcionamiento.");
            Console.WriteLine("");
            Console.WriteLine("Para empezar presione una tecla.");
            Thread.Sleep(500);
            Console.ReadKey();
            Console.Clear();

            char opcion;

            ChromeDriver driver = new ChromeDriver();
            Thread.Sleep(500);
            IWait<IWebDriver> wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30.00));
            Thread.Sleep(500);
            bool inicioSesionNapsis = false;
            bool inicioSesionSige = false;

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


                var archivosTxt = Directory.EnumerateFiles("C:/Alumnos/", "*.txt");

                switch (opcion)
                {
                    case '1':
                        if (!inicioSesionNapsis)
                        {
                            driver.Url = "https://napsis.com/ingresar/";
                            CWInstrucciones();
                            Console.WriteLine(instrucciones[0]);
                            Console.ReadKey();
                            inicioSesionNapsis = true;
                        }
                        
                        driver.Url = "http://ncore.napsis.cl/fichaalumno/ficha-alumno";
                        CWInstrucciones();
                        Console.WriteLine(instrucciones[1]);
                        Console.WriteLine(instrucciones[2]);
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
                                            sw.WriteLine(nombreF);
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
                                    sw.WriteLine(nombre);
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

                        if (archivosTxt.Count() > 0) //if existe más de un archivo en C:/Alumnos
                        {
                            if (!inicioSesionSige)
                            {
                                driver.Url = "https://sige.mineduc.cl/Sige/Login";
                                CWInstrucciones();
                                Console.WriteLine(instrucciones[3]);
                                Console.ReadKey();
                                inicioSesionSige = true;
                            }

                            wait.Until(ExpectedConditions.ElementExists(By.ClassName("ui-icon ui-icon-closethick")));
                            driver.ExecuteScript("document.getElementsByTagName('a')[20].click();");
                            driver.ExecuteScript("document.getElementsByTagName('a')[5].click();");

                            Console.WriteLine("");
                            Console.WriteLine("Este proceso ingresará toda la información almacenada en C:/Alumnos utilizando su cuenta.");
                            Console.WriteLine("Si está seguro de continuar, presione una tecla. Sino, cierre esta ventana.");
                            Console.ReadKey();

                            
                        }
                        else
                        {
                            Console.WriteLine("No existen archivos para copiar a SIGE");
                            Console.WriteLine("Oprima una tecla para volver al menú.");
                            Console.ReadKey();
                        }
                        break;
                    case '3':
                        String[] texto;
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

    }
}