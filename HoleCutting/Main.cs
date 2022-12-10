using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace HoleCutting
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document activetDoc = commandData.Application.ActiveUIDocument.Document;
            Document linedkDoc = activetDoc.Application.Documents.OfType<Document>().Where(d => d.Title.Contains("ИОС")).FirstOrDefault();
            if (linedkDoc == null)
            {
                TaskDialog.Show("Ошибка", "Остутствует модель с указанныи именем");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(activetDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Contains("Отверстие в стене"))
                .FirstOrDefault();

            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено подходящего семейства отверстия");
                return Result.Cancelled;
            }

            ElementClassFilter elementClassFilterDuct = new ElementClassFilter(typeof(Duct));
            ElementClassFilter elementClassFilterPipe = new ElementClassFilter(typeof(Pipe));
            LogicalOrFilter DuctOrPipe = new LogicalOrFilter(elementClassFilterPipe, elementClassFilterDuct);
            List<Element> listElements = new FilteredElementCollector(linedkDoc).WherePasses(DuctOrPipe).ToElements().ToList();
            if (listElements.Count < 1)
            {
                TaskDialog.Show("Ошибка", "Не найдены системы");
                return Result.Cancelled;
            }
            View3D view3D = new FilteredElementCollector(activetDoc)
                .OfClass(typeof(View3D))
                .OfType<View3D>()
                .Where(v => !v.IsTemplate)
                .FirstOrDefault();

            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);
            Transaction transaction = new Transaction(activetDoc, "Расстановка отверский");
            transaction.Start();
            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }
            foreach (Element element in listElements)
            {
                Line line = (element.Location as LocationCurve).Curve as Line;
                XYZ point = line.GetEndPoint(0);
                XYZ direction = line.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= line.Length)
                    .ToList();
                
                foreach (ReferenceWithContext reference in intersections)
                {
                    double proximity = reference.Proximity;
                    Reference refWall = reference.GetReference();
                    Wall wall = activetDoc.GetElement(refWall.ElementId) as Wall;
                    Level level = activetDoc.GetElement(wall.LevelId) as Level;
                    XYZ holePoint = point + (direction * proximity);
                    FamilyInstance hole = activetDoc.Create.NewFamilyInstance(holePoint, familySymbol, wall, level, StructuralType.NonStructural);
                    double widthElement;
                    double heightElement;
                    try
                    {
                        widthElement = (element as MEPCurve).Diameter;
                        heightElement = (element as MEPCurve).Diameter;
                    }
                    catch 
                    {

                        widthElement = (element as MEPCurve).Width;
                        heightElement = (element as MEPCurve).Height;
                    }
                    double k = 1.1;
                    bool width = hole.LookupParameter("Ширина").Set(widthElement*k);
                    bool height = hole.LookupParameter("Высота").Set(heightElement*k);
                }
            }
            transaction.Commit();
            TaskDialog.Show("Результат", "Отверстия прорезаны");
            return Result.Succeeded;
        }
    }
}
