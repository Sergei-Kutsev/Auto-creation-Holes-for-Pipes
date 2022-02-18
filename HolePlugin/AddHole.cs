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

namespace HolePlugin
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class AddHole : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;

            Document ovDoc = arDoc
                .Application
                .Documents //коллекция которую приводим к списку документов, чтобы потом выполнить фильтрацию по имени
                .OfType<Document>()
                .Where(x => x.Title.Contains("ОВ"))
                .FirstOrDefault();

            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ файл");
                return Result.Cancelled;
            }

            Document vkDoc = arDoc
                .Application
                .Documents //коллекция которую приводим к списку документов, чтобы потом выполнить фильтрацию по имени
                .OfType<Document>()
                .Where(x => x.Title.Contains("ВК"))
                .FirstOrDefault();

            if (vkDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ВК файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbol = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстия"))
                .FirstOrDefault();

            if (familySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстия\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(vkDoc)
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
                TaskDialog.Show("Ошибка", "Не найдено семейство 3D вида");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);
            Transaction transaction = new Transaction(arDoc, "Расстановка отверстий");
            transaction.Start();

            if (!familySymbol.IsActive)
                familySymbol.Activate();

            foreach (Duct d in ducts)
            {
                Line curve = (d.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;

                //находим набор всех пересечений в виде объекта ReferenceWithContext
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction) //получаем точку и направление луча
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) //из всех объектов, которые совпадают по заданному критерию, оставляет только один
                    .ToList();
                foreach (ReferenceWithContext refer in intersections) //береберем их и определим точку вставки и добавим туда экземпляр семейства отверстия
                {
                    double proximity = refer.Proximity; //расстояние
                    Reference reference = refer.GetReference(); //у ссылки есть Id
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level; //зная Id из документа можно всегда получить сам объект
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter widthHole = hole.LookupParameter("Ширина");
                    Parameter heightHole = hole.LookupParameter("Высота");
                    widthHole.Set(d.Diameter);
                    heightHole.Set(d.Diameter);
                }
            }
            foreach (Pipe p in pipes)
            {
                Line curve = (p.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;

                //находим набор всех пересечений в виде объекта ReferenceWithContext
                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction) //получаем точку и направление луча
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer()) //из всех объектов, которые совпадают по заданному критерию, оставляет только один
                    .ToList();
                foreach (ReferenceWithContext refer in intersections) //береберем их и определим точку вставки и добавим туда экземпляр семейства отверстия
                {
                    double proximity = refer.Proximity; //расстояние
                    Reference reference = refer.GetReference(); //у ссылки есть Id
                    Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                    Level level = arDoc.GetElement(wall.LevelId) as Level; //зная Id из документа можно всегда получить сам объект
                    XYZ pointHole = point + (direction * proximity);

                    FamilyInstance hole = arDoc.Create.NewFamilyInstance(pointHole, familySymbol, wall, level, StructuralType.NonStructural);
                    Parameter widthHole = hole.LookupParameter("Ширина");
                    Parameter heightHole = hole.LookupParameter("Высота");
                    widthHole.Set(p.Diameter);
                    heightHole.Set(p.Diameter);
                }
            }
            transaction.Commit();
            return Result.Succeeded;
        }
        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y) //будут ли два заданных объекта одинаковыми
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId //LinkedElementId - это Id элемента из связанного файла, они дб одинаковыми
                           && xReference.ElementId == yReference.ElementId; //если у обеих элементов одинаковый ElementId (т.е. точки на одной стене),
                //то вернется trye
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
}
