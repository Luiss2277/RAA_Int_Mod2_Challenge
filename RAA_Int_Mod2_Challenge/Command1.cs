#region Namespaces
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;

#endregion

namespace RAA_Int_Mod2_Challenge
{
    [Transaction(TransactionMode.Manual)]
    public class Command1 : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // this is a variable for the Revit application
            UIApplication uiapp = commandData.Application;

            // this is a variable for the current Revit model
            Document doc = uiapp.ActiveUIDocument.Document;

            //Filtered Element Collector by View
            View curView = doc.ActiveView;
            ViewType curViewType = curView.ViewType;
            FilteredElementCollector collector = new FilteredElementCollector(doc, curView.Id);

            //Category Lists
            List<BuiltInCategory> areaCatList = new List<BuiltInCategory>();
            areaCatList.Add(BuiltInCategory.OST_Areas);

            List<BuiltInCategory> ceilingCatList = new List<BuiltInCategory>();
            ceilingCatList.Add(BuiltInCategory.OST_LightingFixtures);
            ceilingCatList.Add(BuiltInCategory.OST_Rooms);

            List<BuiltInCategory> floorCatList = new List<BuiltInCategory>();
            floorCatList.Add(BuiltInCategory.OST_Walls);
            floorCatList.Add(BuiltInCategory.OST_Doors);
            floorCatList.Add(BuiltInCategory.OST_Furniture);
            floorCatList.Add(BuiltInCategory.OST_Rooms);
            floorCatList.Add(BuiltInCategory.OST_Windows);

            List<BuiltInCategory> sectionCatList = new List<BuiltInCategory>();
            sectionCatList.Add(BuiltInCategory.OST_Rooms);

            //Get Family Symbol names
            FamilySymbol curAreaTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Area Tag")).First();

            FamilySymbol curCurtainWallTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Curtain Wall Tag")).First();

            FamilySymbol curDoorTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Door Tag")).First();

            FamilySymbol curFurnitureTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Furniture Tag")).First();

            FamilySymbol curLightFixTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Lighting Fixture Tag")).First();

            FamilySymbol curRoomTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Room Tag")).First();

            FamilySymbol curWallTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Wall Tag")).First();

            FamilySymbol curWindowTag = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol)).Cast<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("M_Window Tag")).First();

            //Tag Dictionary
            Dictionary<string, FamilySymbol> tags = new Dictionary<string, FamilySymbol>();
            tags.Add("Area", curAreaTag);
            tags.Add("Curtain Walls", curCurtainWallTag);
            tags.Add("Doors", curDoorTag);
            tags.Add("Furniture", curFurnitureTag);
            tags.Add("Lighting Fixtures", curLightFixTag);
            tags.Add("Rooms", curRoomTag);
            tags.Add("Walls", curWallTag);
            tags.Add("Windows", curWindowTag);

            int tagCount = 0;

            using (Transaction t = new Transaction(doc))
            {
                t.Start("Tag, You're It!");

                if (curViewType == ViewType.AreaPlan)
                {
                    //Category Filter
                    ElementMulticategoryFilter areaCatFilter = new ElementMulticategoryFilter(areaCatList);
                    collector.WherePasses(areaCatFilter).WhereElementIsNotElementType();

                    //Tag elements
                    foreach (Element curElem in collector)
                    {
                        XYZ insPoint = GetInsPoint(curElem);
                        if (insPoint == null)
                            continue;

                        ViewPlan curAreaPlan = curView as ViewPlan;
                        Area curArea = curElem as Area;
                        AreaTag currentAreaTag = doc.Create.NewAreaTag(curAreaPlan, curArea, new UV(insPoint.X, insPoint.Y));
                        currentAreaTag.TagHeadPosition = new XYZ(insPoint.X, insPoint.Y, 0);
                        currentAreaTag.HasLeader = false;

                        tagCount++;
                    }
                }
                else if (curViewType == ViewType.CeilingPlan)
                {
                    //Category Filter
                    ElementMulticategoryFilter ceilingCatFilter = new ElementMulticategoryFilter(ceilingCatList);
                    collector.WherePasses(ceilingCatFilter).WhereElementIsNotElementType();

                    //Tag elements
                    foreach (Element curElem in collector)
                    {
                        XYZ insPoint = GetInsPoint(curElem);
                        if (insPoint == null)
                            continue;

                        FamilySymbol curTagType = tags[curElem.Category.Name];
                        Reference curRef = new Reference(curElem);
                        IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);

                        tagCount++;
                    }
                }
                else if (curViewType == ViewType.FloorPlan)
                {
                    //Category Filter
                    ElementMulticategoryFilter floorCatFilter = new ElementMulticategoryFilter(floorCatList);
                    collector.WherePasses(floorCatFilter).WhereElementIsNotElementType();

                    //Tag elements
                    foreach (Element curElem in collector)
                    {
                        XYZ insPoint = GetInsPoint(curElem);
                        if (insPoint == null)
                            continue;
                        //Check for walls and curtain walls
                        if (curElem.Category.Name == "Walls")
                        {
                            Wall curWall = curElem as Wall;
                            WallType curWallType = curWall.WallType;

                            if (curWallType.Kind == WallKind.Curtain)
                            {
                                FamilySymbol curtainTagType = tags["Curtain Walls"];
                                Reference curtainRef = new Reference(curElem);
                                IndependentTag newCurtainTag = IndependentTag.Create(doc, curtainTagType.Id, curView.Id, curtainRef, true, TagOrientation.Horizontal, insPoint);

                                tagCount++;
                            }
                            else
                            {
                                FamilySymbol wallTagType = tags["Walls"];
                                Reference wallRef = new Reference(curElem);
                                IndependentTag newCurtainTag = IndependentTag.Create(doc, wallTagType.Id, curView.Id, wallRef, true, TagOrientation.Horizontal, insPoint);

                                tagCount++;
                            }
                        }
                        else if (curElem.Category.Name == "Windows")
                        {
                            XYZ newInsPoint = new XYZ(insPoint.X, (insPoint.Y + 3), insPoint.Z);

                            FamilySymbol curTagType = tags[curElem.Category.Name];
                            Reference curRef = new Reference(curElem);
                            IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);

                            tagCount++;
                        }
                        else
                        {
                            FamilySymbol curTagType = tags[curElem.Category.Name];
                            Reference curRef = new Reference(curElem);
                            IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id, curRef, false, TagOrientation.Horizontal, insPoint);

                            tagCount++;
                        }
                    }
                }
                else if (curViewType == ViewType.Section)
                {
                    //Category Filter
                    ElementMulticategoryFilter floorCatFilter = new ElementMulticategoryFilter(floorCatList);
                    collector.WherePasses(floorCatFilter).WhereElementIsNotElementType();

                    //Tag elements
                    foreach (Element curElem in collector)
                    {
                        XYZ insPoint = GetInsPoint(curElem);
                        if (insPoint == null)
                            continue;

                        XYZ newInsPoint = new XYZ(insPoint.X, insPoint.Y, (insPoint.Z + 3)); 

                        FamilySymbol curTagType = tags[curElem.Category.Name];
                        Reference curRef = new Reference(curElem);
                        IndependentTag newTag = IndependentTag.Create(doc, curTagType.Id, curView.Id, curRef, false, TagOrientation.Horizontal, newInsPoint);

                        tagCount++;
                    }
                }
                t.Commit();
            }

            TaskDialog.Show("Tag Count", "Elements Tagged: " + tagCount);

            return Result.Succeeded;
        }
        private XYZ GetInsPoint(Element LMN)
        {
            //Get Element Locations
            XYZ insPoint;
            LocationPoint locPoint;
            LocationCurve locCurve;
            Location curLoc = LMN.Location;

            if (curLoc == null)
                return null;

            locPoint = curLoc as LocationPoint;
            if (locPoint != null)
            {
                insPoint = locPoint.Point;
            }
            else
            {
                locCurve = curLoc as LocationCurve;
                Curve curCurve = locCurve.Curve;
                insPoint = GetMidPointBetweenTwoPoints(curCurve.GetEndPoint(0), curCurve.GetEndPoint(1));
            }
            return insPoint;
        }

        private XYZ GetMidPointBetweenTwoPoints(XYZ point1, XYZ point2)
        {
            XYZ midPoint = new XYZ(
                (point1.X + point2.X) / 2,
                (point1.Y + point2.Y) / 2,
                (point1.Z + point2.Z) / 2);
            return midPoint;
        }
        internal static PushButtonData GetButtonData()
        {
            // use this method to define the properties for this command in the Revit ribbon
            string buttonInternalName = "btnCommand1";
            string buttonTitle = "Button 1";

            ButtonDataClass myButtonData1 = new ButtonDataClass(
                buttonInternalName,
                buttonTitle,
                MethodBase.GetCurrentMethod().DeclaringType?.FullName,
                Properties.Resources.Blue_32,
                Properties.Resources.Blue_16,
                "This is a tooltip for Button 1");

            return myButtonData1.Data;
        }
    }
}
