using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Linq;
using System.Net;
using System.Web;
using System.Web.Mvc;
using MVC_CRUD_DiplomandoUASDCodeFirst.Model.DAL;
using MVC_CRUD_DiplomandoUASDCodeFirst.Model.Models;
using System.Data.SqlClient;
using ClosedXML.Excel;
using System.IO;
using System.Xml;

namespace MVC_CRUD_DiplomandoUASDCodeFirst.WEB.Controllers
{
    public class EmpleadosController : Controller
    {

        // Inicio
        private EmpleadoContext db = new EmpleadoContext();

        // GET: Empleadoes
        public ActionResult Index()
        {
            return View(db.Empleados.ToList());
        }

        // GET: Empleadoes/Details/5
        public ActionResult Details(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Empleado empleado = db.Empleados.Find(id);
            if (empleado == null)
            {
                return HttpNotFound();
            }
            return View(empleado);
        }

        // GET: Empleadoes/Create
        public ActionResult Create()
        {
            return View();
        }

        // POST: Empleadoes/Create
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Create([Bind(Include = "EmpleadoID,Nombres,Apellidos,Fecha_Ingreso")] Empleado empleado)
        {
            if (ModelState.IsValid)
            {
                db.Empleados.Add(empleado);
                db.SaveChanges();
                return RedirectToAction("Index");
            }

            return View(empleado);
        }

        // GET: Empleadoes/Edit/5
        public ActionResult Edit(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Empleado empleado = db.Empleados.Find(id);
            if (empleado == null)
            {
                return HttpNotFound();
            }
            return View(empleado);
        }

        // POST: Empleadoes/Edit/5
        // To protect from overposting attacks, please enable the specific properties you want to bind to, for 
        // more details see https://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [ValidateAntiForgeryToken]
        public ActionResult Edit([Bind(Include = "EmpleadoID,Nombres,Apellidos,Fecha_Ingreso")] Empleado empleado)
        {
            if (ModelState.IsValid)
            {
                db.Entry(empleado).State = EntityState.Modified;
                db.SaveChanges();
                return RedirectToAction("Index");
            }
            return View(empleado);
        }

        // GET: Empleadoes/Delete/5
        public ActionResult Delete(int? id)
        {
            if (id == null)
            {
                return new HttpStatusCodeResult(HttpStatusCode.BadRequest);
            }
            Empleado empleado = db.Empleados.Find(id);
            if (empleado == null)
            {
                return HttpNotFound();
            }
            return View(empleado);
        }

        // POST: Empleadoes/Delete/5
        [HttpPost, ActionName("Delete")]
        [ValidateAntiForgeryToken]
        public ActionResult DeleteConfirmed(int id)
        {
            Empleado empleado = db.Empleados.Find(id);
            db.Empleados.Remove(empleado);
            db.SaveChanges();
            return RedirectToAction("Index");
        }


        //Para los primeros 3 Registros de la Tabla

        public FileResult ExportarEmpleadosCSV()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[4] {
                new DataColumn("Codigo"),
                new DataColumn("Nombres"),
                new DataColumn("Apellidos"),
                new DataColumn("Fecha_Ingreso")
            });

            // Tomamos los primeros 3 Registros
            var CurEmpleados = from empleados in db.Empleados.Take(3)
                               select empleados; 
            foreach (var Empleado in CurEmpleados)
            {
                dt.Rows.Add(Empleado.EmpleadoID, Empleado.Nombres, Empleado.Apellidos, Empleado.Fecha_Ingreso);
            }

            using (XLWorkbook wb = new XLWorkbook()) 
            {
                wb.Worksheets.Add(dt);
                using (System.IO.MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Primeros3Empleados.xlsx");
                }
            }
        }

        //Para los primeros 20 Registros de la Tabla

        public FileResult ExportarEmpleadosCSV20()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[4] {
                new DataColumn("Codigo"),
                new DataColumn("Nombres"),
                new DataColumn("Apellidos"),
                new DataColumn("Fecha_Ingreso")
            });

            // Tomamos los primeros 20 Registros
            var CurEmpleados = from empleados in db.Empleados.Take(20)
                               select empleados;
            foreach (var Empleado in CurEmpleados)
            {
                dt.Rows.Add(Empleado.EmpleadoID, Empleado.Nombres, Empleado.Apellidos, Empleado.Fecha_Ingreso);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (System.IO.MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "Primeros20Empleados.xlsx");
                }
            }
        }

        //Para todos los Empleados que su nombre Inicia con J

        public FileResult ExportarEmpleadosCSVInitJ()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[4] {
                new DataColumn("Codigo"),
                new DataColumn("Nombres"),
                new DataColumn("Apellidos"),
                new DataColumn("Fecha_Ingreso")
            });

            var CurEmpleados = from empleados in db.Empleados
                               where empleados.Nombres.StartsWith("J") || empleados.Nombres.StartsWith("j")
                               select empleados ;
            foreach (var Empleado in CurEmpleados)
            {
                dt.Rows.Add(Empleado.EmpleadoID, Empleado.Nombres, Empleado.Apellidos, Empleado.Fecha_Ingreso);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (System.IO.MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EmpleadosNombreEmpiezanJ.xlsx");
                }
            }
        }

        //Para todos los Empleados que su Apellido Contenga "A"

        public FileResult ExportarEmpleadosCSVwithA()
        {
            DataTable dt = new DataTable("Grid");
            dt.Columns.AddRange(new DataColumn[4] {
                new DataColumn("Codigo"),
                new DataColumn("Nombres"),
                new DataColumn("Apellidos"),
                new DataColumn("Fecha_Ingreso")
            });

            var CurEmpleados = from empleados in db.Empleados
                               where empleados.Apellidos.Contains("a") || empleados.Apellidos.Contains("A")
                               select empleados;
            foreach (var Empleado in CurEmpleados)
            {
                dt.Rows.Add(Empleado.EmpleadoID, Empleado.Nombres, Empleado.Apellidos, Empleado.Fecha_Ingreso);
            }

            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt);
                using (System.IO.MemoryStream stream = new MemoryStream())
                {
                    wb.SaveAs(stream);
                    return File(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "EmpleadosApellidosConA.xlsx");
                }
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}
