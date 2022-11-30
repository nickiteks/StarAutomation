from settings import EXCEL_START_ROW
from settings import EXCEL_START_COLL
from settings import GEOM_STP_PATH
from win32com.client import Dispatch
from win32com.client import gencache
import pythoncom

class Methods:

    def numberOfcullums(self,sheet_obj,row_start,coll):
        row_number = 0
        row = row_start
        while True:
            cell_value = sheet_obj.cell(row=row, column=coll).value
            if cell_value == None:
                return row_number
            else:
                row_number = row_number + 1
            
            row = row + 1

    def checkRepeatRow(self,sheet_obj,row,coll):
        start_geometry = []

        for i in range(coll,coll+7):
            start_geometry.append(sheet_obj.cell(row=row, column=i).value)

        for i in reversed(range(EXCEL_START_ROW,row)):
            check_geomentry = [] 
            for j in range(coll,coll+7):
                check_geomentry.append(sheet_obj.cell(row=i, column=j).value)
            
            if start_geometry == check_geomentry:
                return i
            
        return 0

    def changeGeom(self,geomPath,sheet_obj,row):
        module7, api7, const7 = get_kompas_api7()  # Подключаемся к API7
        app7 = api7.Application  # Получаем основной интерфейс
        app7.Visible = True  # Показываем окно пользователю (если скрыто)
        app7.HideMessage = const7.ksHideMessageNo  # Отвечаем НЕТ на любые вопросы программы
        print(app7.ApplicationName(FullName=True))  # Печатаем название программы

        doc7 = app7.Documents.Open(PathName=geomPath,
                                Visible=True,
                                ReadOnly=True)

        #  Подключим константы API Компас
        kompas6_constants = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
        kompas6_constants_3d = gencache.EnsureModule("{2CAF168C-7961-4B90-9DA2-701419BEEFE3}", 0, 1, 0).constants

        #  Подключим описание интерфейсов API5
        kompas6_api5_module = gencache.EnsureModule("{0422828C-F174-495E-AC5D-D31014DBBE87}", 0, 1, 0)
        kompas_object = kompas6_api5_module.KompasObject(
            Dispatch("Kompas.Application.5")._oleobj_.QueryInterface(kompas6_api5_module.KompasObject.CLSID,
                                                                    pythoncom.IID_IDispatch))

        #  Подключим описание интерфейсов API7
        kompas_api7_module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
        application = kompas_api7_module.IApplication(
            Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(kompas_api7_module.IApplication.CLSID,
                                                                    pythoncom.IID_IDispatch))

        Documents = application.Documents
        #  Получим активный документ
        kompas_document = application.ActiveDocument
        kompas_document_3d = kompas_api7_module.IKompasDocument3D(kompas_document)
        iDocument3D = kompas_object.ActiveDocument3D()

        kPart = iDocument3D.GetPart(kompas6_constants_3d.pTop_Part)

        varcoll = kPart.VariableCollection()
        varcoll.refresh()

        values = []

        for i in range(0, 7):
             values.append(sheet_obj.cell(row=row, column = EXCEL_START_COLL + i).value)

        list_collms = ['N1','L1','N2','L2','a2','N3','L3']

        for i in range(len(values)):
            Variable = varcoll.GetByName(list_collms[i], True, True)
            Variable.value = values[i]

        # Перестраиваем модель
        kPart.RebuildModel()
        # Перерисовываем документ
        iDocument3D.RebuildDocument()

        savePath = kompas_document.PathName[:-4]
        print(savePath)

        kompas_document.SaveAs(savePath + f'{row}.stp')
        kompas_document.Close(True)

    def macroGeomCgange(self,row):
        fout = open(f"Macroses/macros{(EXCEL_START_ROW + row)}.java","w")
        fout.write("""
package macro;

import star.amr.AmrModel;
import star.base.neo.DoubleVector;
import star.base.neo.NeoObjectVector;
import star.base.neo.NeoProperty;
import star.base.neo.StringVector;
import star.cadmodeler.*;
import star.combustion.*;
import star.combustion.fgm.FgmCombustionModel;
import star.combustion.fgm.FgmIdealGasModel;
import star.combustion.fgm.FgmReactionModel;
import star.combustion.fgm.FgmTable;
import star.combustion.tablegenerators.*;
import star.common.*;
import star.dualmesher.VolumeControlDualMesherSizeOption;
import star.emissions.NoxModel;
import star.emissions.NoxThreeEquationZeldovichModel;
import star.emissions.ThermalNoxModel;
import star.flow.MassFlowRateProfile;
import star.flow.VelocityMagnitudeProfile;
import star.keturb.KEpsilonTurbulence;
import star.keturb.KeTwoLayerAllYplusWallTreatment;
import star.keturb.RkeTwoLayerTurbModel;
import star.material.MaterialDataBaseManager;
import star.material.MultiComponentGasModel;
import star.meshing.*;
import star.reactions.ReactingModel;
import star.resurfacer.VolumeControlResurfacerSizeOption;
import star.segregatedenergy.SegregatedFluidEnthalpyModel;
import star.segregatedflow.SegregatedFlowModel;
import star.turbulence.RansTurbulenceModel;
import star.turbulence.TurbulentModel;
import star.vis.*;

import javax.swing.*;

public class %s extends StarMacro {

    public void execute() {

        importGeometry();
    }

    private void importGeometry() {

        Simulation simulation =
                getActiveSimulation();

        CadModel cadModel =
                simulation.get(SolidModelManager.class).createSolidModel();

        cadModel.resetSystemOptions();

        ImportCadFileFeature importCadFileFeature =
                cadModel.importCadFile(resolvePath("%s"), true, false, false, false, false, false, false, true, false, true, NeoProperty.fromString("{\'CGR\': 0, \'IFC\': 0, \'STEP\': 0, \'NX\': 0, \'CATIAV5\': 0, \'SE\': 0, \'JT\': 0}"), false);

        star.cadmodeler.Body cadmodelerBody_0 =
                ((star.cadmodeler.Body) importCadFileFeature.getBodyByIndex(1));

        Face face_0 =
                ((Face) importCadFileFeature.getFaceByLocation(cadmodelerBody_0, new DoubleVector(new double[]{0.01, -0.22786658652431685, 0.10277042815333508})));

        cadModel.setFaceNameAttributes(new NeoObjectVector(new Object[]{face_0}), "Air blades", false);

        Face face_1 =
                ((Face) importCadFileFeature.getFaceByLocation(cadmodelerBody_0, new DoubleVector(new double[]{0.01, -0.23814856654623856, 0.4957704212230306})));

        cadModel.setFaceNameAttributes(new NeoObjectVector(new Object[]{face_1}), "Air input", false);

        Face face_2 =
                ((Face) importCadFileFeature.getFaceByLocation(cadmodelerBody_0, new DoubleVector(new double[]{0.01, -0.04266834375631555, -0.13363422057600013})));

        cadModel.setFaceNameAttributes(new NeoObjectVector(new Object[]{face_2}), "CH4", false);

        Face face_3 =
                ((Face) importCadFileFeature.getFaceByLocation(cadmodelerBody_0, new DoubleVector(new double[]{0.01, -0.042284638094857586, 0.08802686971235965})));

        cadModel.setFaceNameAttributes(new NeoObjectVector(new Object[]{face_3}), "No Gas 1", false);

        Face face_4 =
                ((Face) importCadFileFeature.getFaceByLocation(cadmodelerBody_0, new DoubleVector(new double[]{0.01, -0.009175941885004056, 0.004994952138007335})));

        cadModel.setFaceNameAttributes(new NeoObjectVector(new Object[]{face_4}), "No Gas 2", false);

        Face face_5 =
                ((Face) importCadFileFeature.getFaceByLocation(cadmodelerBody_0, new DoubleVector(new double[]{-10.667, 0.6180132, -0.6180132})));

        cadModel.setFaceNameAttributes(new NeoObjectVector(new Object[]{face_5}), "Outlet", false);

        cadmodelerBody_0.getUnNamedFacesDefaultAttributeName();

        simulation.get(SolidModelManager.class).endEditCadModel(cadModel);

        cadModel.createParts(new NeoObjectVector(new Object[]{cadmodelerBody_0}), new NeoObjectVector(new Object[]{}), true, false, 1, false, false, 3, "SharpEdges", 30.0, 2, true, 1.0E-5, false);
    }

}       
         """ % (f"macros{(EXCEL_START_ROW + row)}",f"{GEOM_STP_PATH}{(EXCEL_START_ROW + row)}.stp"))

def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const