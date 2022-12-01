// STAR-CCM+ macro: SetupExhaustPost.java
package macro;

import java.util.*;

import star.common.*;
import star.base.neo.*;
import star.vis.*;
import star.base.report.*;
import star.flow.*;

public class SetupManifoldPost extends StarMacro {

    private Simulation simulation_0;
    private Units units_0;
    private String unitsName = "C";
    private String fieldFunction1 = "Temperature";
    private String fieldFunction2 = "WallYplus";

    public void execute() {
        execute0();
    }

    private void execute0() {

        simulation_0 = getActiveSimulation();

        units_0 = ((Units) simulation_0.getUnitsManager().getObject(unitsName));
        PrimitiveFieldFunction temperature = ((PrimitiveFieldFunction) simulation_0.getFieldFunctionManager().getFunction(fieldFunction1));
        PrimitiveFieldFunction wallyplus = ((PrimitiveFieldFunction) simulation_0.getFieldFunctionManager().getFunction(fieldFunction2));

        String regionNames[] = getRegionNames();

        if (regionNames == null) {
            simulation_0.println("### ERROR: Leaving macro.");
        } else {

            Region fluid = simulation_0.getRegionManager().getRegion(regionNames[0]);
            Region solid = simulation_0.getRegionManager().getRegion(regionNames[1]);

            String[] boundaryNames = getBoundaryNames(fluid, solid);
            if (boundaryNames == null) {
                simulation_0.println("### ERROR: Leaving macro.");
            } else {
                Boundary outlet = fluid.getBoundaryManager().getBoundary(boundaryNames[0]);
                simulation_0.println("# Starting Post Setup");

                DirectBoundaryInterfaceBoundary interf
                        = ((DirectBoundaryInterfaceBoundary) fluid.getBoundaryManager().getBoundary(boundaryNames[1]));
                Boundary wall = solid.getBoundaryManager().getBoundary(boundaryNames[2]);

                // Create a Mass Averaged Report for Fluid and Solid
                // private void massAveragedReport(Region fluid, Region solid, PrimitiveFieldFunction pff)
                massAveragedReport(fluid, solid, temperature);

                // Create Mass Flow Averaged Report for Outlet
                // private void massFlowAveragedReport(Boundary bnd, PrimitiveFieldFunction pff)
                massFlowAveragedReport(outlet, temperature);

                // Create 3 scalar scenes
                // 1. Temperature on Solid Wall
                // 2. Temperature on Fluid Interface
                // 3. Wall Y+ on Fluid Interface
                // private void createScalarScenes(Region fluid, Boundary bnd, DirectBoundaryInterfaceBoundary interf,
                //    PrimitiveFieldFunction pff1, PrimitiveFieldFunction pff2)
                createScalarScenes(fluid, wall, interf, temperature, wallyplus);

                // Create mesh scene on interface
                // private void createMesh2Scene(DirectBoundaryInterfaceBoundary interf)
                createMesh2Scene(interf);

                simulation_0.getSimulationIterator().run();
            }
        }
    }

    private String[] getRegionNames() {

        Collection<Region> colr = simulation_0.getRegionManager().getRegions();
        // Default expected names. Corrected if necessary in the loop.
        String[] names = new String[2];
        names[0] = "";
        names[1] = "";

        if (colr.size() != 2) {
            simulation_0.println("### Only two regions are allowed or expected:");
            simulation_0.println("  ### Number of regions: " + colr.size());
            names = null;
            return names;
        } else {
            for (Region reg : colr) {
                if (reg.getRegionType().getPresentationName().equalsIgnoreCase("Fluid Region")
                        || reg.getRegionType().getPresentationName().equalsIgnoreCase("Porous Region")) {
                    names[0] = reg.getPresentationName();
                    simulation_0.println("# Name of fluid: " + reg.getPresentationName());
                } else if (reg.getRegionType().getPresentationName().equalsIgnoreCase("Solid Region")) {
                    names[1] = reg.getPresentationName();
                    simulation_0.println("# Name of solid: " + reg.getPresentationName());
                }
            }
            if (names[0].equals("") || names[1].equals("")) {
                simulation_0.println("### Fluid or solid region is missing:");
                simulation_0.println("  ### Fluid region: " + names[0] + "\n  ### Solid region: " + names[1]);
                names = null;
                return names;
            } else {
                return names;
            }
        }
    }

    private String[] getBoundaryNames(Region fluid, Region solid) {

        String[] names = new String[3];
        // Default expected names. Corrected if necessary in the loop.
        names[0] = "";
        names[1] = "";
        names[2] = "";

        Collection<Boundary> colbf = fluid.getBoundaryManager().getBoundaries();

        for (Boundary bnd : colbf) {
            if (bnd.getBoundaryType().getPresentationName().equalsIgnoreCase("Pressure Outlet")) {
                simulation_0.println("# Name of outlet: " + bnd.getPresentationName());
                names[0] = bnd.getPresentationName();
            } else if (bnd.getPresentationName().contains("[")) {
                simulation_0.println("# Name of interface: " + bnd.getPresentationName());
                names[1] = bnd.getPresentationName();
            } else {
                simulation_0.println("### not used or wrong type: " + bnd.getPresentationName());
            }
        }

        Collection<Boundary> colbs = solid.getBoundaryManager().getBoundaries();
        // Default expected names. Corrected if necessary in the loop.

        if (colbs.size() != 2) {
            simulation_0.println("### Two boundaries are expected. It is not "
                    + "possible to find the outer wall:");
            simulation_0.println("  ### Number of solid boundaries: " + colbs.size());
            names = null;
            return names;
        } else {
            for (Boundary bnd : colbs) {
                if (bnd.getPresentationName().contains("[")) {
                } else {
                    names[2] = bnd.getPresentationName();
                    simulation_0.println("## Name of solid outer wall: " + names[2]);
                }
            }
        }

        if (names[0].equals("") || names[1].equals("") | names[2].equals("")) {
            simulation_0.println("### One or more boundaries are missing:");
            simulation_0.println("  ### outlet: " + names[0] + "\n  ### interface: " + names[1] + "\n  ### solid wall: " + names[2]);
            names = null;
            return names;
        } else {
            return names;
        }

    }

    private void massAveragedReport(Region fluid, Region solid, PrimitiveFieldFunction pff) {

        String fluidReportName = "Averaged " + fluid.getPresentationName() + " " + pff.getPresentationName();
        String solidReportName = "Averaged " + solid.getPresentationName() + " " + pff.getPresentationName();

        MassAverageReport massAverageReport_0
                = simulation_0.getReportManager().createReport(MassAverageReport.class);
        massAverageReport_0.setPresentationName(fluidReportName);
        massAverageReport_0.setScalar(pff);
        massAverageReport_0.getParts().setObjects(fluid);
        massAverageReport_0.setUnits(units_0);

        MassAverageReport massAverageReport_1
                = simulation_0.getReportManager().createReport(MassAverageReport.class);
        massAverageReport_1.setPresentationName(solidReportName);
        massAverageReport_1.copyProperties(massAverageReport_0);
        massAverageReport_1.getParts().setObjects(solid);
        massAverageReport_0.setUnits(units_0);

        simulation_0.getMonitorManager().createMonitorAndPlot(new NeoObjectVector(new Object[]{massAverageReport_0, massAverageReport_1}), true, "Reports Plot");
        ReportMonitor reportMonitor_0
                = ((ReportMonitor) simulation_0.getMonitorManager().getMonitor(fluidReportName + " Monitor"));
        ReportMonitor reportMonitor_1
                = ((ReportMonitor) simulation_0.getMonitorManager().getMonitor(solidReportName + " Monitor"));
        MonitorPlot monitorPlot_0
                = simulation_0.getPlotManager().createMonitorPlot(new NeoObjectVector(new Object[]{reportMonitor_0, reportMonitor_1}), "Reports Plot");
        monitorPlot_0.open();

    }

    private void massFlowAveragedReport(Boundary bnd, PrimitiveFieldFunction pff) {

        String reportName = "Bulk " + bnd.getPresentationName() + " " + pff.getPresentationName();

        MassFlowAverageReport massFlowAverageReport_0
                = simulation_0.getReportManager().createReport(MassFlowAverageReport.class);
        massFlowAverageReport_0.setPresentationName(reportName);
        massFlowAverageReport_0.setScalar(pff);
        massFlowAverageReport_0.getParts().setObjects(bnd);
        massFlowAverageReport_0.setUnits(units_0);

        simulation_0.getMonitorManager().createMonitorAndPlot(new NeoObjectVector(new Object[]{massFlowAverageReport_0}), true, "%1$s Plot");
        ReportMonitor reportMonitor_0 = ((ReportMonitor) simulation_0.getMonitorManager().getMonitor(reportName + " Monitor"));
        MonitorPlot monitorPlot_0 = simulation_0.getPlotManager().createMonitorPlot(new NeoObjectVector(new Object[]{reportMonitor_0}), reportName + " Plot");
        monitorPlot_0.open();
        
    }

    private void createScalarScenes(Region fluid, Boundary bnd, DirectBoundaryInterfaceBoundary interf,
            PrimitiveFieldFunction pff1, PrimitiveFieldFunction pff2) {

        // Scalar Scene for temperature on soid wall
        Scene scene_3 = simulation_0.getSceneManager().createScene("Temperature");
        scene_3.setPresentationName("Temperature");
        ScalarDisplayer scalarDisplayer_0 = scene_3.getDisplayerManager().createScalarDisplayer("Scalar");
        scene_3.open(true);

        CurrentView currentView_3 = scene_3.getCurrentView();
        currentView_3.setInput(new DoubleVector(new double[]{-0.019769827474829048, 0.08452465191935596, 0.03470089789311794}), new DoubleVector(new double[]{-0.4448332968171574, -0.08118675869461253, -0.6964823744294228}), new DoubleVector(new double[]{-0.05130473184957331, 0.9799987999459978, -0.19227630273709706}), 0.22498538953528152, 0);

        scalarDisplayer_0.getParts().setObjects(bnd);
        scalarDisplayer_0.setFillMode(1);
        scalarDisplayer_0.getScalarDisplayQuantity().setFieldFunction(pff1);
        scalarDisplayer_0.getScalarDisplayQuantity().setUnits(units_0);

        // Copy of Scene 3, Scalar Scene for temperature on interface
        Scene scene_4 = simulation_0.getSceneManager().createScene("Copy of Temperature");
        scene_4.setPresentationName("Interface Temperature");
        scene_4.copyProperties(scene_3);
        ScalarDisplayer scalarDisplayer_1 = ((ScalarDisplayer) scene_4.getDisplayerManager().getDisplayer("Scalar"));
        scalarDisplayer_1.getParts().setObjects(interf);
        scene_4.open(true);

        // Copy of Scene 4, Scalar Scene for Wall y+ on interface
        Scene scene_5 = simulation_0.getSceneManager().createScene("Copy of Interface Temperature");
        scene_5.setPresentationName("Wall Y+");
        scene_5.copyProperties(scene_4);
        ScalarDisplayer scalarDisplayer_2 = ((ScalarDisplayer) scene_5.getDisplayerManager().getDisplayer("Scalar"));
        scalarDisplayer_2.initialize();
        scalarDisplayer_2.getScalarDisplayQuantity().setFieldFunction(pff2);
        scene_5.open(true);

    }

    private void createMesh2Scene(DirectBoundaryInterfaceBoundary interf) {

        Scene scene_6 = simulation_0.getSceneManager().createScene("Mesh Scene Interface");

        PartDisplayer partDisplayer_21 = scene_6.getDisplayerManager().createPartDisplayer("Geometry", -1, 4);
        partDisplayer_21.setPresentationName("Mesh 1");
        partDisplayer_21.initialize();
        partDisplayer_21.setOutline(false);
        partDisplayer_21.setSurface(true);
        partDisplayer_21.setMesh(true);

        scene_6.open(true);

        CurrentView currentView_0 = scene_6.getCurrentView();
        currentView_0.setInput(new DoubleVector(new double[]{-0.030843563450813294, 0.08020754438228082, 0.08350931107997894}), new DoubleVector(new double[]{-0.030843563450813294, 0.08020754438228082, 1.0220580247935627}), new DoubleVector(new double[]{0.0, 1.0, 0.0}), 0.24501037962812278, 0);

        partDisplayer_21.getParts().setObjects(interf);

    }
}
