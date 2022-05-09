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
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;

namespace Gusev_Revit_PlagIn_Practicum_Final
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class CreationModel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;
            Autodesk.Revit.Creation.Document document = doc.Create;
            string levelNumber;

            // Если в модели нет параметра "Номер помещения" добавляем его в проект к категории OST_Rooms(Помещения)
            CategorySet categorySet = new CategorySet();
            categorySet.Insert(Category.GetCategory(doc, BuiltInCategory.OST_Rooms));
            using (Transaction transaction3 = new Transaction(doc, "Добавление парамектра Номер помещения"))
            {
                transaction3.Start();
                bool isParameterOk=CreateSharedParameter(uiapp.Application, doc, "Номер этажа",
                     categorySet, BuiltInParameterGroup.PG_IDENTITY_DATA, true);
                if (!isParameterOk)
                    return Result.Cancelled;
                transaction3.Commit();
            }

            // Если в модели не добавлено семейство "Марка помещения1" сообщить пользователю и выйти
            FamilySymbol usedTagFamilySymbol = new FilteredElementCollector(doc)
                                                   .OfClass(typeof(FamilySymbol))
                                                   .OfCategory(BuiltInCategory.OST_RoomTags)
                                                   .OfType<FamilySymbol>()
                                                   .Where(x => x.Name.Equals("Марка помещения1"))
                                                   .FirstOrDefault();
            if (usedTagFamilySymbol == null)
            {
                TaskDialog.Show("Ошибка", "Не загружено семейство \"Марка помещения1\".");
                return Result.Cancelled;
            }
            // Сделать семейство "Марка помещения1" семейством по умолчанию для марок помещений
            using (Transaction transaction4 = new Transaction(doc,
                   "установка семейства по умолчанию для марки помещений"))
            {
                transaction4.Start();
                ElementId usedTagFamilySymbolId = usedTagFamilySymbol.Id;
                ElementId defaultRoomTagId = new ElementId(BuiltInCategory.OST_RoomTags);
                doc.SetDefaultFamilyTypeId(defaultRoomTagId, usedTagFamilySymbolId);
                transaction4.Commit();
            }

            // уничтожаем все существующие комнаты
            List<Room> rooms = GetAllRooms(doc);
            List<ElementId> roomsId = new List<ElementId>();

            // Получаем существующие уровни в проекте
            List<Level> levels = GetLevels(doc);
            //rooms[0].Number
/*            List<Room> roomsLevel_1;
            foreach (Level level in levels)
            {
                roomsLevel_1 = GetRooms(doc, level);
                if (roomsLevel_1.Count == 0)
                    continue;
                else
                {
                    Parameter room1Number = roomsLevel_1[0].get_Parameter(Autodesk.Revit.DB.BuiltInParameter.ROOM_NUMBER);
                    room1Number.Set(0);
                    break;
                }
            }
*/            

            foreach (Room room in rooms)
                roomsId.Add(room.Id);

            Transaction transaction1 = new Transaction(doc, "Удаление существующих комнат");
            transaction1.Start();
            doc.Delete(roomsId);
            rooms.Clear();
            roomsId.Clear();
            transaction1.Commit();

            // Создаем новые комнаты
            // Получаем фазу текущего вида
            Parameter para = doc.ActiveView.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.VIEW_PHASE);
            Autodesk.Revit.DB.ElementId phaseId = para.AsElementId();
            Phase m_defaultPhase = commandData.Application.ActiveUIDocument.Document.GetElement(phaseId) as Phase;

            if (levels == null)
            {
                TaskDialog.Show("Ошибка", "В модели не ни одного уровня");
                return Result.Cancelled;
            }

            Autodesk.Revit.DB.View initialView = uiDoc.ActiveView;

            foreach (Level level in levels)
            {
                Autodesk.Revit.DB.View view = doc.GetElement(level.FindAssociatedPlanViewId())
                                              as Autodesk.Revit.DB.View;
                if ((view != null)&& (view.ViewType == Autodesk.Revit.DB.ViewType.FloorPlan))
                {
                    uiDoc.ActiveView = view;
                    para = doc.ActiveView.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.VIEW_PHASE);
                    phaseId = para.AsElementId();
                    m_defaultPhase = doc.GetElement(phaseId) as Phase;
                    if (m_defaultPhase == null)
                    {
                        Autodesk.Revit.UI.TaskDialog.Show("Revit", "The phase of the active view is null, you can't create spaces in a null phase");
                        return Result.Cancelled;
                    }

                    Transaction transaction = new Transaction(doc, "Создание комнат и меток комнат");
                    transaction.Start();

                    // Создание комнат
                    roomsId=document.NewRooms2(level, m_defaultPhase).ToList();
                    foreach (ElementId id in roomsId)
                        rooms.Add(doc.GetElement(id) as Room);

                    transaction.Commit();
                    levelNumber = "";
                    string[] subs = level.Name.Split();
                    if ((subs[0].Equals("Этаж"))|| (subs[0].Equals("Уровень"))|| (subs[0].Equals("Level")))
                        levelNumber = subs[1];

                    List <RoomTag> roomTags = GetRoomsTags(doc, level);

                    Transaction transaction2 = new Transaction(doc, "установка параметра метки Номер этажа");
                    transaction2.Start();

                    foreach (Room room in rooms)
                    {
                        Parameter levelNumberPar = room.LookupParameter("Номер этажа");
                        levelNumberPar.Set(Convert.ToInt16(levelNumber));
                    }
                    rooms.Clear();
                    transaction2.Commit();
                }
                
                /*
                List<Autodesk.Revit.DB.View> levelViews = GetViewsAssociatedWithLevel(doc, level);

                foreach (Autodesk.Revit.DB.View view in levelViews)
                {
                    try
                    {
                        if (view.ViewType == Autodesk.Revit.DB.ViewType.FloorPlan)
                        {
                            uiDoc.ActiveView = view;
                            para = doc.ActiveView.get_Parameter(Autodesk.Revit.DB.BuiltInParameter.VIEW_PHASE);
                            phaseId = para.AsElementId();
                            m_defaultPhase = doc.GetElement(phaseId) as Phase;
                            if (m_defaultPhase == null)
                            {
                                Autodesk.Revit.UI.TaskDialog.Show("Revit", "The phase of the active view is null, you can't create spaces in a null phase");
                                return Result.Cancelled;
                            }
                            int a = 0;
                            document.NewRooms2(level, m_defaultPhase);
                            break;
                        }
                        else
                        {
                            continue;
                        }
                    }
                    catch (Exception ex)
                    {
                        Autodesk.Revit.UI.TaskDialog.Show("Revit", ex.Message);
                    }
                }
                */

              // Room room = new Room();
                

            }
            uiDoc.ActiveView = initialView;
    


            return Result.Succeeded;
        }

        private bool CreateSharedParameter(Autodesk.Revit.ApplicationServices.Application application, 
                                           Document doc,
                                           string parameterName,
                                           CategorySet categorySet,
                                           BuiltInParameterGroup builtInParameterGroup,
                                           bool isInstance)
        {
            DefinitionFile definitionFile = application.OpenSharedParameterFile();
            if (definitionFile == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл общих параметров");
                return false;
            }
            Definition definition = definitionFile.Groups
                                                  .SelectMany(group => group.Definitions)
                                                  .FirstOrDefault(def=>def.Name.Equals(parameterName));
            if (definition == null)
            {
                TaskDialog.Show("Ошибка", "Не найден параметр: "+ parameterName);
                return false;
            }

            Autodesk.Revit.DB.Binding binding = application.Create.NewTypeBinding(categorySet);
            if (isInstance)
                binding = application.Create.NewInstanceBinding(categorySet);

            BindingMap map = doc.ParameterBindings;
            if (map.Contains(definition))
                return true;
            else 
                map.Insert(definition, binding, builtInParameterGroup);
            return true;
        }



        // Метод получающий существующие уровни из документа, имена которых перечислены в входном списке

        public List<Level> GetLevels(Document doc, List<String> levelNames)
        {
            List<Level> listNamedlevel = new List<Level>();
            List<Level> listlevel = new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .OfType<Level>()
                            .ToList();
            foreach (String levelName in levelNames)
            {
                try
                {
                    listNamedlevel.Add(listlevel.FirstOrDefault(x => x.Name.Equals(levelName)));
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }
            return listNamedlevel;
        }

        public List<Level> GetLevels(Document doc)
        {
            List<Level> listlevel = new FilteredElementCollector(doc)
                            .OfClass(typeof(Level))
                            .OfType<Level>()
                            .ToList();
            return listlevel;
        }

        public List<Room> GetRooms(Document doc, Level level)
        {
            List<Room> listRooms = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Rooms)
                            .OfType<Room>()
                            .Where(x=>x.Level==level)
                            .ToList();
            return listRooms;
        }

        public List<Room> GetRooms(Document doc)
        {
            List<Room> listRooms = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Rooms)
                            .OfType<Room>()
                            .ToList();
            return listRooms;
        }

        public List<Autodesk.Revit.DB.View> GetViewsAssociatedWithLevel(Document doc, Level level)
        {
            List<Autodesk.Revit.DB.View> listViews = new FilteredElementCollector(doc)
                            .OfClass(typeof(Autodesk.Revit.DB.View))
                            .OfType<Autodesk.Revit.DB.View>()
                            .Where(x=>x.LevelId==level.Id)
                            .ToList();
            return listViews;
        }

        public List<Room> GetAllRooms(Document doc)
        {
            List<Room> listRooms = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_Rooms)
                            .OfType<Room>()
                            .ToList();
            return listRooms;
        }

        public List<RoomTag> GetRoomsTags(Document doc, Level level)
        {
            List<RoomTag> listRoomsTags = new FilteredElementCollector(doc)
                            .OfCategory(BuiltInCategory.OST_RoomTags)
                            .OfType<RoomTag>()
                            .Where(x => x.Room.Level == level)
                            .ToList();
            return listRoomsTags;
        }

    }
}
