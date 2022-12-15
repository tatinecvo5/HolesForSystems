using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;

namespace HolesForSystems
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uiArDoc = commandData.Application.ActiveUIDocument;
            Document arDoc = uiArDoc.Document;
            Document ovDoc = arDoc.Application.Documents.OfType<Document>()
                             .Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл ОВ");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                                            .OfClass(typeof(FamilySymbol))
                                            .OfCategory(BuiltInCategory.OST_GenericModel)
                                            .OfType<FamilySymbol>()
                                            .Where(x => x.FamilyName.Equals("Отверстие прямоугольное"))
                                            .FirstOrDefault();
            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстие прямоугольное\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                                   .OfClass(typeof(Duct))
                                   .OfType<Duct>()
                                   .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
                       .OfClass(typeof(Pipe))
                       .OfType<Pipe>()
                       .ToList();

            View3D view3D = new FilteredElementCollector(arDoc)
                                .OfClass(typeof(View3D))
                                .OfType<View3D>()
                                .Where(x => !x.IsTemplate)
                                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D-вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector
                = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)),
                                           FindReferenceTarget.Element, view3D);
            Transaction transaction = new Transaction(arDoc);
            transaction.Start("Активация семейства");

            if (!familySymbol.IsActive)
                familySymbol.Activate();
            transaction.Commit();

            Transaction transaction1 = new Transaction(arDoc);
            transaction1.Start("Расстановка отверстий");

            foreach (Duct d in ducts)
            {
                Line curve = (d.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                                           .Where(x => x.Proximity <= curve.Length)
                                           .Distinct(new ReferenceWithContextElementEqualityComparer())
                                           .ToList();
                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole,
                                   familySymbol, wall, level,
                                   Autodesk.Revit.DB.Structure.
                                   StructuralType.NonStructural);
                    Parameter width = hole.LookupParameter("ширина");
                    Parameter height = hole.LookupParameter("высота");
                    width.Set(d.Diameter + 1);
                    height.Set(d.Diameter + 1);
                    
                }
            }

            foreach (Pipe p in pipes)
            {
                Line curveP = (p.Location as LocationCurve).Curve as Line;
                XYZ pointP = curveP.GetEndPoint(0);
                XYZ directionP = curveP.Direction;

                List<ReferenceWithContext> intersectionsP = referenceIntersector.Find(pointP, directionP)
                                           .Where(x => x.Proximity <= curveP.Length)
                                           .Distinct(new ReferenceWithContextElementEqualityComparer())
                                           .ToList();
                foreach (ReferenceWithContext referP in intersectionsP)
                {
                    double proximityP = referP.Proximity;
                    Reference referenceP = referP.GetReference();
                    Wall wall = arDoc.GetElement(referenceP.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level;
                    XYZ pointHoleP = pointP + (directionP * proximityP);

                    FamilyInstance holeP = arDoc.Create.NewFamilyInstance(pointHoleP,
                                   familySymbol, wall, level,
                                   Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                    Parameter width = holeP.LookupParameter("ширина");
                    Parameter height = holeP.LookupParameter("высота");
                    
                        width.Set(p.Diameter + 1);
                        height.Set(p.Diameter + 1);
                    
                }
            }

            transaction1.Commit();
            return Result.Succeeded;
        }
    }

    public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, x)) return false;
            if (ReferenceEquals(null, y)) return false;

            var xReference = x.GetReference();

            var yReference = y.GetReference();

            return xReference.LinkedElementId == yReference.LinkedElementId
                   && xReference.ElementId == yReference.ElementId;
        }

        public int GetHashCode(ReferenceWithContext obj)
        {
            var reference = obj.GetReference();

            unchecked
            {
                return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
            }
        }
    }
}
