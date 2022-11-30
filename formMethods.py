from settings import EXCEL_START_ROW
from settings import EXCEL_START_COLL
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

    def macroGeomCgange(self,row,geom):
        fout = open(f"Macroses/macros{row}.java","w")
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
        createCylinderParts();
        createVolumeMeshControl();
        createBoundaries();
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

    private void createCylinderParts() {

        Simulation simulation =
                getActiveSimulation();

        Units units_0 =
                simulation.getUnitsManager().getPreferredUnits(Dimensions.Builder().length(1).build());

        MeshPartFactory meshPartFactory_0 =
                simulation.get(MeshPartFactory.class);

        SimpleCylinderPart simpleCylinderPart_0 =
                meshPartFactory_0.createNewCylinderPart(simulation.get(SimulationPartManager.class));

        simpleCylinderPart_0.setDoNotRetessellate(true);

        LabCoordinateSystem labCoordinateSystem_0 =
                simulation.getCoordinateSystemManager().getLabCoordinateSystem();

        simpleCylinderPart_0.setCoordinateSystem(labCoordinateSystem_0);

        simpleCylinderPart_0.getStartCoordinate().setCoordinateSystem(labCoordinateSystem_0);

        simpleCylinderPart_0.getStartCoordinate().setCoordinate(units_0, units_0, units_0, new DoubleVector(new double[]{0.01, 0.0, 0.0}));

        simpleCylinderPart_0.getEndCoordinate().setCoordinateSystem(labCoordinateSystem_0);

        simpleCylinderPart_0.getEndCoordinate().setCoordinate(units_0, units_0, units_0, new DoubleVector(new double[]{-0.667, 0.0, 0.0}));

        simpleCylinderPart_0.getRadius().setUnits(units_0);

        simpleCylinderPart_0.getRadius().setValue(0.721);

        simpleCylinderPart_0.getTessellationDensityOption().setSelected(TessellationDensityOption.Type.MEDIUM);

        simpleCylinderPart_0.rebuildSimpleShapePart();

        simpleCylinderPart_0.setDoNotRetessellate(false);

        SimpleCylinderPart simpleCylinderPart_1 =
                meshPartFactory_0.createNewCylinderPart(simulation.get(SimulationPartManager.class));

        simpleCylinderPart_1.setDoNotRetessellate(true);

        simpleCylinderPart_1.setCoordinateSystem(labCoordinateSystem_0);

        simpleCylinderPart_1.getStartCoordinate().setCoordinateSystem(labCoordinateSystem_0);

        simpleCylinderPart_1.getStartCoordinate().setCoordinate(units_0, units_0, units_0, new DoubleVector(new double[]{-0.667, 0.0, 0.0}));

        simpleCylinderPart_1.getEndCoordinate().setCoordinateSystem(labCoordinateSystem_0);

        simpleCylinderPart_1.getEndCoordinate().setCoordinate(units_0, units_0, units_0, new DoubleVector(new double[]{-1.111, 0.0, 0.0}));

        simpleCylinderPart_1.getRadius().setUnits(units_0);

        simpleCylinderPart_1.getRadius().setValue(0.535);

        simpleCylinderPart_1.getTessellationDensityOption().setSelected(TessellationDensityOption.Type.MEDIUM);

        simpleCylinderPart_1.rebuildSimpleShapePart();

        simpleCylinderPart_1.setDoNotRetessellate(false);

        SimpleCylinderPart simpleCylinderPart_2 =
                meshPartFactory_0.createNewCylinderPart(simulation.get(SimulationPartManager.class));

        simpleCylinderPart_2.setDoNotRetessellate(true);

        simpleCylinderPart_2.setCoordinateSystem(labCoordinateSystem_0);

        simpleCylinderPart_2.getStartCoordinate().setCoordinateSystem(labCoordinateSystem_0);

        simpleCylinderPart_2.getStartCoordinate().setCoordinate(units_0, units_0, units_0, new DoubleVector(new double[]{-1.111, 0.0, 0.0}));

        simpleCylinderPart_2.getEndCoordinate().setCoordinateSystem(labCoordinateSystem_0);

        simpleCylinderPart_2.getEndCoordinate().setCoordinate(units_0, units_0, units_0, new DoubleVector(new double[]{-10.667, 0.0, 0.0}));

        simpleCylinderPart_2.getRadius().setUnits(units_0);

        simpleCylinderPart_2.getRadius().setValue(1.5);

        simpleCylinderPart_2.getTessellationDensityOption().setSelected(TessellationDensityOption.Type.MEDIUM);

        simpleCylinderPart_2.rebuildSimpleShapePart();

        simpleCylinderPart_2.setDoNotRetessellate(false);
    }

    private void createVolumeMeshControl(){

        Simulation simulation =
                getActiveSimulation();

        Units units_0 =
                simulation.getUnitsManager().getPreferredUnits(Dimensions.Builder().length(1).build());

        SimpleCylinderPart simpleCylinderPart_0 =
                ((SimpleCylinderPart) simulation.get(SimulationPartManager.class).getPart("Cylinder"));

        SimpleCylinderPart simpleCylinderPart_1 =
                ((SimpleCylinderPart) simulation.get(SimulationPartManager.class).getPart("Cylinder 2"));

        SimpleCylinderPart simpleCylinderPart_2 =
                ((SimpleCylinderPart) simulation.get(SimulationPartManager.class).getPart("Cylinder 3"));

        SolidModelPart solidModelPart_0 =
                ((SolidModelPart) simulation.get(SimulationPartManager.class).getPart("Body 1_Body"));

        AutoMeshOperation autoMeshOperation_0 =
                simulation.get(MeshOperationManager.class).createAutoMeshOperation(new StringVector(new String[]{"star.resurfacer.ResurfacerAutoMesher", "star.dualmesher.DualAutoMesher"}), new NeoObjectVector(new Object[]{solidModelPart_0}));

        VolumeCustomMeshControl volumeCustomMeshControl_0 =
                autoMeshOperation_0.getCustomMeshControls().createVolumeControl();

        VolumeCustomMeshControl volumeCustomMeshControl_1 =
                autoMeshOperation_0.getCustomMeshControls().createVolumeControl();

        VolumeCustomMeshControl volumeCustomMeshControl_2 =
                autoMeshOperation_0.getCustomMeshControls().createVolumeControl();

        SurfaceCustomMeshControl surfaceCustomMeshControl_0 =
                autoMeshOperation_0.getCustomMeshControls().createSurfaceControl();

        volumeCustomMeshControl_0.getGeometryObjects().setQuery(null);

        volumeCustomMeshControl_0.getGeometryObjects().setObjects(simpleCylinderPart_0);

        volumeCustomMeshControl_1.getGeometryObjects().setQuery(null);

        volumeCustomMeshControl_1.getGeometryObjects().setObjects(simpleCylinderPart_1);

        volumeCustomMeshControl_2.getGeometryObjects().setQuery(null);

        volumeCustomMeshControl_2.getGeometryObjects().setObjects(simpleCylinderPart_2);

        surfaceCustomMeshControl_0.getGeometryObjects().setQuery(null);

        PartSurface partSurface_0 =
                ((PartSurface) solidModelPart_0.getPartSurfaceManager().getPartSurface("Default"));

        PartSurface partSurface_1 =
                ((PartSurface) solidModelPart_0.getPartSurfaceManager().getPartSurface("Outlet"));

        surfaceCustomMeshControl_0.getGeometryObjects().setObjects(partSurface_0, partSurface_1);

        surfaceCustomMeshControl_0.getCustomConditions().get(PartsTargetSurfaceSizeOption.class).setSelected(PartsTargetSurfaceSizeOption.Type.CUSTOM);

        PartsTargetSurfaceSize partsTargetSurfaceSize_0 =
                surfaceCustomMeshControl_0.getCustomValues().get(PartsTargetSurfaceSize.class);

        Units units_1 =
                ((Units) simulation.getUnitsManager().getObject(""));

        partsTargetSurfaceSize_0.getRelativeSizeScalar().setValueAndUnits(20.0, units_1);

        VolumeControlResurfacerSizeOption volumeControlResurfacerSizeOption_0 =
                volumeCustomMeshControl_0.getCustomConditions().get(VolumeControlResurfacerSizeOption.class);

        volumeControlResurfacerSizeOption_0.setVolumeControlBaseSizeOption(true);

        VolumeControlDualMesherSizeOption volumeControlDualMesherSizeOption_0 =
                volumeCustomMeshControl_0.getCustomConditions().get(VolumeControlDualMesherSizeOption.class);

        volumeControlDualMesherSizeOption_0.setVolumeControlBaseSizeOption(true);

        VolumeControlSize volumeControlSize_0 =
                volumeCustomMeshControl_0.getCustomValues().get(VolumeControlSize.class);

        volumeControlSize_0.getRelativeSizeScalar().setValueAndUnits(2.0, units_1);

        VolumeControlResurfacerSizeOption volumeControlResurfacerSizeOption_1 =
                volumeCustomMeshControl_1.getCustomConditions().get(VolumeControlResurfacerSizeOption.class);

        volumeControlResurfacerSizeOption_1.setVolumeControlBaseSizeOption(true);

        VolumeControlDualMesherSizeOption volumeControlDualMesherSizeOption_1 =
                volumeCustomMeshControl_1.getCustomConditions().get(VolumeControlDualMesherSizeOption.class);

        volumeControlDualMesherSizeOption_1.setVolumeControlBaseSizeOption(true);

        VolumeControlSize volumeControlSize_1 =
                volumeCustomMeshControl_1.getCustomValues().get(VolumeControlSize.class);

        volumeControlSize_1.getRelativeSizeScalar().setValueAndUnits(2.0, units_1);

        VolumeControlResurfacerSizeOption volumeControlResurfacerSizeOption_2 =
                volumeCustomMeshControl_2.getCustomConditions().get(VolumeControlResurfacerSizeOption.class);

        volumeControlResurfacerSizeOption_2.setVolumeControlBaseSizeOption(true);

        VolumeControlDualMesherSizeOption volumeControlDualMesherSizeOption_2 =
                volumeCustomMeshControl_2.getCustomConditions().get(VolumeControlDualMesherSizeOption.class);

        volumeControlDualMesherSizeOption_2.setVolumeControlBaseSizeOption(true);

        VolumeControlSize volumeControlSize_2 =
                volumeCustomMeshControl_2.getCustomValues().get(VolumeControlSize.class);

        volumeControlSize_2.getRelativeSizeScalar().setValueAndUnits(5.0, units_1);

        autoMeshOperation_0.getDefaultValues().get(BaseSize.class).setValueAndUnits(0.5, units_0);

    }

    private void createBoundaries() {

        Simulation simulation =
                getActiveSimulation();

        SolidModelPart solidModelPart =
                ((SolidModelPart) simulation.get(SimulationPartManager.class).getPart("Body 1_Body"));

        simulation.getRegionManager().newRegionsFromParts(new NeoObjectVector(new Object[]{solidModelPart}), "OneRegionPerPart", null, "OneBoundaryPerPartSurface", null, "OneFeatureCurve", null, RegionManager.CreateInterfaceMode.BOUNDARY, "OneEdgeBoundaryPerPart", null);

        Region region_0 =
                simulation.getRegionManager().getRegion("Body 1_Body");

        Boundary boundary_0 =
                region_0.getBoundaryManager().getBoundary("Air blades");

        InletBoundary inletBoundary_0 =
                ((InletBoundary) simulation.get(ConditionTypeManager.class).get(InletBoundary.class));

        boundary_0.setBoundaryType(inletBoundary_0);

        Boundary boundary_1 =
                region_0.getBoundaryManager().getBoundary("Air input");

        boundary_1.setBoundaryType(inletBoundary_0);

        Boundary boundary_2 =
                region_0.getBoundaryManager().getBoundary("CH4");

        MassFlowBoundary massFlowBoundary_0 =
                ((MassFlowBoundary) simulation.get(ConditionTypeManager.class).get(MassFlowBoundary.class));

        boundary_2.setBoundaryType(massFlowBoundary_0);

        Boundary boundary_3 =
                region_0.getBoundaryManager().getBoundary("Outlet");

        PressureBoundary pressureBoundary_0 =
                ((PressureBoundary) simulation.get(ConditionTypeManager.class).get(PressureBoundary.class));

        boundary_3.setBoundaryType(pressureBoundary_0);

    }

}       
         """ % (f"macros{row}",f"{geom}{row}.stp"))

def get_kompas_api7():
    module = gencache.EnsureModule("{69AC2981-37C0-4379-84FD-5DD2F3C0A520}", 0, 1, 0)
    api = module.IKompasAPIObject(
        Dispatch("Kompas.Application.7")._oleobj_.QueryInterface(module.IKompasAPIObject.CLSID,
                                                                 pythoncom.IID_IDispatch))
    const = gencache.EnsureModule("{75C9F5D0-B5B8-4526-8681-9903C567D2ED}", 0, 1, 0).constants
    return module, api, const