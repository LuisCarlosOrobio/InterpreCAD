using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace CADInterpreter
{
    // Basic 3D point structure
    public struct Point3D
    {
        public double X, Y, Z;
        public Point3D(double x, double y, double z) { X = x; Y = y; Z = z; }
        public override string ToString() => $"({X}, {Y}, {Z})";
    }

    // Vector structure
    public struct Vector3D
    {
        public double X, Y, Z;
        public Vector3D(double x, double y, double z) { X = x; Y = y; Z = z; }
        public override string ToString() => $"({X}, {Y}, {Z})";
    }

    // Color structure
    public struct Color3D
    {
        public int R, G, B;
        public Color3D(int r, int g, int b) { R = r; G = g; B = b; }
        public override string ToString() => $"RGB({R}, {G}, {B})";
    }

    // Simple command structure
    public class Command
    {
        public string Type { get; set; }
        public Dictionary<string, object> Parameters { get; set; } = new Dictionary<string, object>();
        public int LineNumber { get; set; }
    }

    // Exception for parsing errors
    public class InterpreterException : Exception
    {
        public int LineNumber { get; }
        public InterpreterException(string message, int lineNumber = 0) : base(message) 
        { 
            LineNumber = lineNumber; 
        }
    }

    // Complete CAD interpreter class with all AutoCAD entities
    public class CADInterpreter
    {
        private List<Command> commands = new List<Command>();
        private StringBuilder autoCADCode = new StringBuilder();
        private int currentLine = 0;
        private Dictionary<string, string> variables = new Dictionary<string, string>();
        
        public void ParseScript(string script)
        {
            var lines = script.Split('\n');
            commands.Clear();
            currentLine = 0;

            foreach (var line in lines)
            {
                currentLine++;
                try
                {
                    ParseCommand(line);
                }
                catch (Exception ex)
                {
                    throw new InterpreterException($"Line {currentLine}: {ex.Message}", currentLine);
                }
            }
        }

        public void ParseCommand(string input)
        {
            input = input.Trim();
            if (string.IsNullOrEmpty(input) || input.StartsWith("//"))
                return;

            var command = new Command { LineNumber = currentLine };
            
            // Handle variable assignment: VAR name = value
            if (input.StartsWith("VAR "))
            {
                ParseVariable(input);
                return;
            }

            // Simple regex to parse: COMMAND_NAME(param1=value1, param2=value2)
            var match = Regex.Match(input, @"(\w+)\s*\((.*)\)");
            
            if (!match.Success)
            {
                throw new InterpreterException($"Invalid command format: {input}");
            }

            command.Type = match.Groups[1].Value.ToUpper();
            string paramString = match.Groups[2].Value;

            // Parse parameters
            if (!string.IsNullOrEmpty(paramString))
            {
                ParseParameters(paramString, command.Parameters);
            }

            commands.Add(command);
            Console.WriteLine($"Parsed: {command.Type} with {command.Parameters.Count} parameters");
        }

        private void ParseVariable(string input)
        {
            // VAR name = value
            var match = Regex.Match(input, @"VAR\s+(\w+)\s*=\s*(.+)");
            if (match.Success)
            {
                string name = match.Groups[1].Value;
                string value = match.Groups[2].Value.Trim();
                variables[name] = value;
                Console.WriteLine($"Variable {name} = {value}");
            }
        }

        private void ParseParameters(string paramString, Dictionary<string, object> parameters)
        {
            var parts = SplitParameters(paramString);
            
            foreach (var part in parts)
            {
                var keyValue = part.Split('=');
                if (keyValue.Length != 2) 
                    throw new InterpreterException($"Invalid parameter format: {part}");

                string key = keyValue[0].Trim();
                string value = keyValue[1].Trim();

                // Substitute variables
                value = SubstituteVariables(value);
                parameters[key] = ParseValue(value);
            }
        }

        private string SubstituteVariables(string value)
        {
            foreach (var variable in variables)
            {
                value = value.Replace($"${variable.Key}", variable.Value);
            }
            return value;
        }

        private List<string> SplitParameters(string input)
        {
            var result = new List<string>();
            var current = "";
            int bracketCount = 0;
            bool inQuotes = false;

            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
                
                if (c == '"' && (i == 0 || input[i-1] != '\\'))
                    inQuotes = !inQuotes;
                else if (!inQuotes)
                {
                    if (c == '[' || c == '(') bracketCount++;
                    else if (c == ']' || c == ')') bracketCount--;
                    else if (c == ',' && bracketCount == 0)
                    {
                        result.Add(current.Trim());
                        current = "";
                        continue;
                    }
                }
                
                current += c;
            }
            
            if (!string.IsNullOrEmpty(current.Trim()))
                result.Add(current.Trim());
                
            return result;
        }

        private object ParseValue(string value)
        {
            value = value.Trim();

            // Parse point [x,y,z]
            if (value.StartsWith("[") && value.EndsWith("]"))
            {
                var coords = value.Substring(1, value.Length - 2)
                    .Split(',')
                    .Select(s => double.Parse(s.Trim()))
                    .ToArray();
                
                if (coords.Length == 3)
                    return new Point3D(coords[0], coords[1], coords[2]);
                else
                    throw new InterpreterException($"Point must have 3 coordinates: {value}");
            }

            // Parse vector <x,y,z>
            if (value.StartsWith("<") && value.EndsWith(">"))
            {
                var coords = value.Substring(1, value.Length - 2)
                    .Split(',')
                    .Select(s => double.Parse(s.Trim()))
                    .ToArray();
                
                if (coords.Length == 3)
                    return new Vector3D(coords[0], coords[1], coords[2]);
                else
                    throw new InterpreterException($"Vector must have 3 coordinates: {value}");
            }

            // Parse color RGB(r,g,b)
            var colorMatch = Regex.Match(value, @"RGB\s*\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)");
            if (colorMatch.Success)
            {
                int r = int.Parse(colorMatch.Groups[1].Value);
                int g = int.Parse(colorMatch.Groups[2].Value);
                int b = int.Parse(colorMatch.Groups[3].Value);
                return new Color3D(r, g, b);
            }

            // Parse string
            if (value.StartsWith("\"") && value.EndsWith("\""))
                return value.Substring(1, value.Length - 2);

            // Parse boolean
            if (value.ToLower() == "true") return true;
            if (value.ToLower() == "false") return false;

            // Parse number
            if (double.TryParse(value, out double numValue))
                return numValue;

            // Default to string
            return value;
        }

        public string GenerateAutoCADCode()
        {
            autoCADCode.Clear();
            
            // Add header
            autoCADCode.AppendLine("using Autodesk.AutoCAD.ApplicationServices;");
            autoCADCode.AppendLine("using Autodesk.AutoCAD.DatabaseServices;");
            autoCADCode.AppendLine("using Autodesk.AutoCAD.EditorInput;");
            autoCADCode.AppendLine("using Autodesk.AutoCAD.Geometry;");
            autoCADCode.AppendLine("using Autodesk.AutoCAD.Runtime;");
            autoCADCode.AppendLine("using Autodesk.AutoCAD.Colors;");
            autoCADCode.AppendLine();
            autoCADCode.AppendLine("namespace GeneratedCAD");
            autoCADCode.AppendLine("{");
            autoCADCode.AppendLine("    public class CADCommands");
            autoCADCode.AppendLine("    {");
            autoCADCode.AppendLine("        [CommandMethod(\"RUNGENERATED\")]");
            autoCADCode.AppendLine("        public void RunGenerated()");
            autoCADCode.AppendLine("        {");
            autoCADCode.AppendLine("            Document doc = Application.DocumentManager.MdiActiveDocument;");
            autoCADCode.AppendLine("            Database db = doc.Database;");
            autoCADCode.AppendLine("            Editor ed = doc.Editor;");
            autoCADCode.AppendLine();
            autoCADCode.AppendLine("            using (Transaction tr = db.TransactionManager.StartTransaction())");
            autoCADCode.AppendLine("            {");
            autoCADCode.AppendLine("                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);");
            autoCADCode.AppendLine("                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);");
            autoCADCode.AppendLine();

            // Generate commands
            foreach (var cmd in commands)
            {
                GenerateCommandCode(cmd);
            }

            // Add footer
            autoCADCode.AppendLine("                tr.Commit();");
            autoCADCode.AppendLine("            }");
            autoCADCode.AppendLine("        }");
            autoCADCode.AppendLine("    }");
            autoCADCode.AppendLine("}");

            return autoCADCode.ToString();
        }

        private void GenerateCommandCode(Command cmd)
        {
            autoCADCode.AppendLine($"                // {cmd.Type} command");
            
            switch (cmd.Type)
            {
                // 3D Solids
                case "BOX":
                    GenerateBoxCode(cmd.Parameters);
                    break;
                case "SPHERE":
                    GenerateSphereCode(cmd.Parameters);
                    break;
                case "CYLINDER":
                    GenerateCylinderCode(cmd.Parameters);
                    break;
                case "CONE":
                    GenerateConeCode(cmd.Parameters);
                    break;
                case "TORUS":
                    GenerateTorusCode(cmd.Parameters);
                    break;
                case "WEDGE":
                    GenerateWedgeCode(cmd.Parameters);
                    break;
                case "PYRAMID":
                    GeneratePyramidCode(cmd.Parameters);
                    break;

                // Curves and Lines
                case "LINE":
                    GenerateLineCode(cmd.Parameters);
                    break;
                case "CIRCLE":
                    GenerateCircleCode(cmd.Parameters);
                    break;
                case "ARC":
                    GenerateArcCode(cmd.Parameters);
                    break;
                case "ELLIPSE":
                    GenerateEllipseCode(cmd.Parameters);
                    break;
                case "SPLINE":
                    GenerateSplineCode(cmd.Parameters);
                    break;
                case "POLYLINE":
                    GeneratePolylineCode(cmd.Parameters);
                    break;
                case "POLYLINE3D":
                    GeneratePolyline3DCode(cmd.Parameters);
                    break;
                case "RAY":
                    GenerateRayCode(cmd.Parameters);
                    break;
                case "XLINE":
                    GenerateXlineCode(cmd.Parameters);
                    break;

                // Text and Annotations
                case "TEXT":
                    GenerateTextCode(cmd.Parameters);
                    break;
                case "MTEXT":
                    GenerateMTextCode(cmd.Parameters);
                    break;

                // Dimensions
                case "DIMENSION_LINEAR":
                    GenerateLinearDimensionCode(cmd.Parameters);
                    break;
                case "DIMENSION_ANGULAR":
                    GenerateAngularDimensionCode(cmd.Parameters);
                    break;
                case "DIMENSION_RADIAL":
                    GenerateRadialDimensionCode(cmd.Parameters);
                    break;
                case "DIMENSION_DIAMETRIC":
                    GenerateDiametricDimensionCode(cmd.Parameters);
                    break;

                // Surfaces and Regions
                case "HATCH":
                    GenerateHatchCode(cmd.Parameters);
                    break;
                case "REGION":
                    GenerateRegionCode(cmd.Parameters);
                    break;
                case "SURFACE":
                    GenerateSurfaceCode(cmd.Parameters);
                    break;
                case "MESH":
                    GenerateMeshCode(cmd.Parameters);
                    break;

                // Points and Blocks
                case "POINT":
                    GeneratePointCode(cmd.Parameters);
                    break;
                case "BLOCK_REF":
                    GenerateBlockRefCode(cmd.Parameters);
                    break;

                // Organization
                case "LAYER":
                    GenerateLayerCode(cmd.Parameters);
                    break;
                case "GROUP":
                    GenerateGroupCode(cmd.Parameters);
                    break;

                // Transformations
                case "MOVE":
                    GenerateMoveCode(cmd.Parameters);
                    break;
                case "ROTATE":
                    GenerateRotateCode(cmd.Parameters);
                    break;
                case "SCALE":
                    GenerateScaleCode(cmd.Parameters);
                    break;
                case "MIRROR":
                    GenerateMirrorCode(cmd.Parameters);
                    break;
                case "ARRAY_RECTANGULAR":
                    GenerateArrayRectangularCode(cmd.Parameters);
                    break;
                case "ARRAY_POLAR":
                    GenerateArrayPolarCode(cmd.Parameters);
                    break;

                // Boolean Operations
                case "UNION":
                    GenerateUnionCode(cmd.Parameters);
                    break;
                case "SUBTRACT":
                    GenerateSubtractCode(cmd.Parameters);
                    break;
                case "INTERSECT":
                    GenerateIntersectCode(cmd.Parameters);
                    break;

                // Materials and Visualization
                case "MATERIAL":
                    GenerateMaterialCode(cmd.Parameters);
                    break;
                case "LIGHT":
                    GenerateLightCode(cmd.Parameters);
                    break;
                case "CAMERA":
                    GenerateCameraCode(cmd.Parameters);
                    break;

                // Advanced Entities
                case "LEADER":
                    GenerateLeaderCode(cmd.Parameters);
                    break;
                case "MLEADER":
                    GenerateMLeaderCode(cmd.Parameters);
                    break;
                case "MLINE":
                    GenerateMlineCode(cmd.Parameters);
                    break;
                case "TRACE":
                    GenerateTraceCode(cmd.Parameters);
                    break;
                case "SOLID":
                    GenerateSolid2DCode(cmd.Parameters);
                    break;
                case "FACE":
                    GenerateFaceCode(cmd.Parameters);
                    break;

                default:
                    autoCADCode.AppendLine($"                // Unknown command: {cmd.Type}");
                    break;
            }
            autoCADCode.AppendLine();
        }

        // 3D Solids implementation
        private void GenerateBoxCode(Dictionary<string, object> parameters)
        {
            var width = GetParameter<double>(parameters, "width", 1.0);
            var height = GetParameter<double>(parameters, "height", 1.0);
            var depth = GetParameter<double>(parameters, "depth", 1.0);
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                Solid3d box = new Solid3d();");
            autoCADCode.AppendLine($"                box.CreateBox({width}, {height}, {depth});");
            autoCADCode.AppendLine($"                box.TransformBy(Matrix3d.Displacement(new Vector3d({position.X}, {position.Y}, {position.Z})));");
            autoCADCode.AppendLine($"                btr.AppendEntity(box);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(box, true);");
        }

        private void GenerateSphereCode(Dictionary<string, object> parameters)
        {
            var radius = GetParameter<double>(parameters, "radius", 1.0);
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                Solid3d sphere = new Solid3d();");
            autoCADCode.AppendLine($"                sphere.CreateSphere({radius});");
            autoCADCode.AppendLine($"                sphere.TransformBy(Matrix3d.Displacement(new Vector3d({position.X}, {position.Y}, {position.Z})));");
            autoCADCode.AppendLine($"                btr.AppendEntity(sphere);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(sphere, true);");
        }

        private void GenerateCylinderCode(Dictionary<string, object> parameters)
        {
            var radius = GetParameter<double>(parameters, "radius", 1.0);
            var height = GetParameter<double>(parameters, "height", 1.0);
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                Solid3d cylinder = new Solid3d();");
            autoCADCode.AppendLine($"                cylinder.CreateFrustum({height}, {radius}, {radius}, {radius});");
            autoCADCode.AppendLine($"                cylinder.TransformBy(Matrix3d.Displacement(new Vector3d({position.X}, {position.Y}, {position.Z})));");
            autoCADCode.AppendLine($"                btr.AppendEntity(cylinder);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(cylinder, true);");
        }

        private void GenerateConeCode(Dictionary<string, object> parameters)
        {
            var radius = GetParameter<double>(parameters, "radius", 1.0);
            var height = GetParameter<double>(parameters, "height", 1.0);
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                Solid3d cone = new Solid3d();");
            autoCADCode.AppendLine($"                cone.CreateFrustum({height}, {radius}, {radius}, 0);");
            autoCADCode.AppendLine($"                cone.TransformBy(Matrix3d.Displacement(new Vector3d({position.X}, {position.Y}, {position.Z})));");
            autoCADCode.AppendLine($"                btr.AppendEntity(cone);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(cone, true);");
        }

        private void GenerateTorusCode(Dictionary<string, object> parameters)
        {
            var majorRadius = GetParameter<double>(parameters, "majorRadius", 2.0);
            var minorRadius = GetParameter<double>(parameters, "minorRadius", 0.5);
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                Solid3d torus = new Solid3d();");
            autoCADCode.AppendLine($"                torus.CreateTorus({majorRadius}, {minorRadius});");
            autoCADCode.AppendLine($"                torus.TransformBy(Matrix3d.Displacement(new Vector3d({position.X}, {position.Y}, {position.Z})));");
            autoCADCode.AppendLine($"                btr.AppendEntity(torus);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(torus, true);");
        }

        private void GenerateWedgeCode(Dictionary<string, object> parameters)
        {
            var length = GetParameter<double>(parameters, "length", 2.0);
            var width = GetParameter<double>(parameters, "width", 1.0);
            var height = GetParameter<double>(parameters, "height", 1.0);
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                Solid3d wedge = new Solid3d();");
            autoCADCode.AppendLine($"                wedge.CreateWedge({length}, {width}, {height});");
            autoCADCode.AppendLine($"                wedge.TransformBy(Matrix3d.Displacement(new Vector3d({position.X}, {position.Y}, {position.Z})));");
            autoCADCode.AppendLine($"                btr.AppendEntity(wedge);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(wedge, true);");
        }

        private void GeneratePyramidCode(Dictionary<string, object> parameters)
        {
            var radius = GetParameter<double>(parameters, "radius", 1.0);
            var height = GetParameter<double>(parameters, "height", 2.0);
            var sides = GetParameter<double>(parameters, "sides", 4);
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                Solid3d pyramid = new Solid3d();");
            autoCADCode.AppendLine($"                pyramid.CreatePyramid({height}, (int){sides}, {radius}, 0);");
            autoCADCode.AppendLine($"                pyramid.TransformBy(Matrix3d.Displacement(new Vector3d({position.X}, {position.Y}, {position.Z})));");
            autoCADCode.AppendLine($"                btr.AppendEntity(pyramid);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(pyramid, true);");
        }

        // Curves and Lines implementation
        private void GenerateLineCode(Dictionary<string, object> parameters)
        {
            var start = GetParameter<Point3D>(parameters, "start", new Point3D(0, 0, 0));
            var end = GetParameter<Point3D>(parameters, "end", new Point3D(1, 0, 0));

            autoCADCode.AppendLine($"                Line line = new Line();");
            autoCADCode.AppendLine($"                line.StartPoint = new Point3d({start.X}, {start.Y}, {start.Z});");
            autoCADCode.AppendLine($"                line.EndPoint = new Point3d({end.X}, {end.Y}, {end.Z});");
            autoCADCode.AppendLine($"                btr.AppendEntity(line);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(line, true);");
        }

        private void GenerateCircleCode(Dictionary<string, object> parameters)
        {
            var radius = GetParameter<double>(parameters, "radius", 1.0);
            var center = GetParameter<Point3D>(parameters, "center", new Point3D(0, 0, 0));
            var normal = GetParameter<Vector3D>(parameters, "normal", new Vector3D(0, 0, 1));

            autoCADCode.AppendLine($"                Circle circle = new Circle();");
            autoCADCode.AppendLine($"                circle.Center = new Point3d({center.X}, {center.Y}, {center.Z});");
            autoCADCode.AppendLine($"                circle.Radius = {radius};");
            autoCADCode.AppendLine($"                circle.Normal = new Vector3d({normal.X}, {normal.Y}, {normal.Z});");
            autoCADCode.AppendLine($"                btr.AppendEntity(circle);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(circle, true);");
        }

        private void GenerateArcCode(Dictionary<string, object> parameters)
        {
            var radius = GetParameter<double>(parameters, "radius", 1.0);
            var center = GetParameter<Point3D>(parameters, "center", new Point3D(0, 0, 0));
            var startAngle = GetParameter<double>(parameters, "startAngle", 0.0) * Math.PI / 180.0;
            var endAngle = GetParameter<double>(parameters, "endAngle", 90.0) * Math.PI / 180.0;
            var normal = GetParameter<Vector3D>(parameters, "normal", new Vector3D(0, 0, 1));

            autoCADCode.AppendLine($"                Arc arc = new Arc();");
            autoCADCode.AppendLine($"                arc.Center = new Point3d({center.X}, {center.Y}, {center.Z});");
            autoCADCode.AppendLine($"                arc.Radius = {radius};");
            autoCADCode.AppendLine($"                arc.StartAngle = {startAngle};");
            autoCADCode.AppendLine($"                arc.EndAngle = {endAngle};");
            autoCADCode.AppendLine($"                arc.Normal = new Vector3d({normal.X}, {normal.Y}, {normal.Z});");
            autoCADCode.AppendLine($"                btr.AppendEntity(arc);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(arc, true);");
        }

        private void GenerateEllipseCode(Dictionary<string, object> parameters)
        {
            var center = GetParameter<Point3D>(parameters, "center", new Point3D(0, 0, 0));
            var majorAxis = GetParameter<Vector3D>(parameters, "majorAxis", new Vector3D(2, 0, 0));
            var radiusRatio = GetParameter<double>(parameters, "radiusRatio", 0.5);
            var normal = GetParameter<Vector3D>(parameters, "normal", new Vector3D(0, 0, 1));

            autoCADCode.AppendLine($"                Ellipse ellipse = new Ellipse();");
            autoCADCode.AppendLine($"                ellipse.Center = new Point3d({center.X}, {center.Y}, {center.Z});");
            autoCADCode.AppendLine($"                ellipse.MajorAxis = new Vector3d({majorAxis.X}, {majorAxis.Y}, {majorAxis.Z});");
            autoCADCode.AppendLine($"                ellipse.RadiusRatio = {radiusRatio};");
            autoCADCode.AppendLine($"                ellipse.Normal = new Vector3d({normal.X}, {normal.Y}, {normal.Z});");
            autoCADCode.AppendLine($"                btr.AppendEntity(ellipse);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(ellipse, true);");
        }

        private void GenerateSplineCode(Dictionary<string, object> parameters)
        {
            var start = GetParameter<Point3D>(parameters, "start", new Point3D(0, 0, 0));
            var end = GetParameter<Point3D>(parameters, "end", new Point3D(5, 5, 0));
            var control1 = GetParameter<Point3D>(parameters, "control1", new Point3D(1, 2, 0));
            var control2 = GetParameter<Point3D>(parameters, "control2", new Point3D(4, 3, 0));

            autoCADCode.AppendLine($"                Point3dCollection fitPoints = new Point3dCollection();");
            autoCADCode.AppendLine($"                fitPoints.Add(new Point3d({start.X}, {start.Y}, {start.Z}));");
            autoCADCode.AppendLine($"                fitPoints.Add(new Point3d({control1.X}, {control1.Y}, {control1.Z}));");
            autoCADCode.AppendLine($"                fitPoints.Add(new Point3d({control2.X}, {control2.Y}, {control2.Z}));");
            autoCADCode.AppendLine($"                fitPoints.Add(new Point3d({end.X}, {end.Y}, {end.Z}));");
            autoCADCode.AppendLine($"                Spline spline = new Spline(fitPoints, 3, 0.0);");
            autoCADCode.AppendLine($"                btr.AppendEntity(spline);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(spline, true);");
        }

        private void GeneratePolylineCode(Dictionary<string, object> parameters)
        {
            var start = GetParameter<Point3D>(parameters, "start", new Point3D(0, 0, 0));
            var end = GetParameter<Point3D>(parameters, "end", new Point3D(1, 1, 0));
            var width = GetParameter<double>(parameters, "width", 0.0);

            autoCADCode.AppendLine($"                Polyline pline = new Polyline();");
            autoCADCode.AppendLine($"                pline.AddVertexAt(0, new Point2d({start.X}, {start.Y}), 0, {width}, {width});");
            autoCADCode.AppendLine($"                pline.AddVertexAt(1, new Point2d({end.X}, {end.Y}), 0, {width}, {width});");
            autoCADCode.AppendLine($"                btr.AppendEntity(pline);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(pline, true);");
        }

        private void GeneratePolyline3DCode(Dictionary<string, object> parameters)
        {
            var start = GetParameter<Point3D>(parameters, "start", new Point3D(0, 0, 0));
            var end = GetParameter<Point3D>(parameters, "end", new Point3D(1, 1, 1));

            autoCADCode.AppendLine($"                Polyline3d pline3d = new Polyline3d();");
            autoCADCode.AppendLine($"                btr.AppendEntity(pline3d);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(pline3d, true);");
            autoCADCode.AppendLine($"                PolylineVertex3d vertex1 = new PolylineVertex3d(new Point3d({start.X}, {start.Y}, {start.Z}));");
            autoCADCode.AppendLine($"                PolylineVertex3d vertex2 = new PolylineVertex3d(new Point3d({end.X}, {end.Y}, {end.Z}));");
            autoCADCode.AppendLine($"                pline3d.AppendVertex(vertex1);");
            autoCADCode.AppendLine($"                pline3d.AppendVertex(vertex2);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(vertex1, true);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(vertex2, true);");
        }

        private void GenerateRayCode(Dictionary<string, object> parameters)
        {
            var basePoint = GetParameter<Point3D>(parameters, "basePoint", new Point3D(0, 0, 0));
            var direction = GetParameter<Vector3D>(parameters, "direction", new Vector3D(1, 0, 0));

            autoCADCode.AppendLine($"                Ray ray = new Ray();");
            autoCADCode.AppendLine($"                ray.BasePoint = new Point3d({basePoint.X}, {basePoint.Y}, {basePoint.Z});");
            autoCADCode.AppendLine($"                ray.UnitDir = new Vector3d({direction.X}, {direction.Y}, {direction.Z}).GetNormal();");
            autoCADCode.AppendLine($"                btr.AppendEntity(ray);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(ray, true);");
        }

        private void GenerateXlineCode(Dictionary<string, object> parameters)
        {
            var basePoint = GetParameter<Point3D>(parameters, "basePoint", new Point3D(0, 0, 0));
            var direction = GetParameter<Vector3D>(parameters, "direction", new Vector3D(1, 0, 0));

            autoCADCode.AppendLine($"                Xline xline = new Xline();");
            autoCADCode.AppendLine($"                xline.BasePoint = new Point3d({basePoint.X}, {basePoint.Y}, {basePoint.Z});");
            autoCADCode.AppendLine($"                xline.UnitDir = new Vector3d({direction.X}, {direction.Y}, {direction.Z}).GetNormal();");
            autoCADCode.AppendLine($"                btr.AppendEntity(xline);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(xline, true);");
        }

        // Text and Annotations implementation
        private void GenerateTextCode(Dictionary<string, object> parameters)
        {
            var text = GetParameter<string>(parameters, "text", "Sample Text");
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));
            var height = GetParameter<double>(parameters, "height", 1.0);
            var rotation = GetParameter<double>(parameters, "rotation", 0.0) * Math.PI / 180.0;

            autoCADCode.AppendLine($"                DBText dbText = new DBText();");
            autoCADCode.AppendLine($"                dbText.Position = new Point3d({position.X}, {position.Y}, {position.Z});");
            autoCADCode.AppendLine($"                dbText.Height = {height};");
            autoCADCode.AppendLine($"                dbText.TextString = \"{text}\";");
            autoCADCode.AppendLine($"                dbText.Rotation = {rotation};");
            autoCADCode.AppendLine($"                btr.AppendEntity(dbText);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(dbText, true);");
        }

        private void GenerateMTextCode(Dictionary<string, object> parameters)
        {
            var text = GetParameter<string>(parameters, "text", "Sample MText");
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));
            var height = GetParameter<double>(parameters, "height", 1.0);
            var width = GetParameter<double>(parameters, "width", 10.0);

            autoCADCode.AppendLine($"                MText mtext = new MText();");
            autoCADCode.AppendLine($"                mtext.Location = new Point3d({position.X}, {position.Y}, {position.Z});");
            autoCADCode.AppendLine($"                mtext.TextHeight = {height};");
            autoCADCode.AppendLine($"                mtext.Width = {width};");
            autoCADCode.AppendLine($"                mtext.Contents = \"{text}\";");
            autoCADCode.AppendLine($"                btr.AppendEntity(mtext);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(mtext, true);");
        }

        // Dimensions implementation
        private void GenerateLinearDimensionCode(Dictionary<string, object> parameters)
        {
            var point1 = GetParameter<Point3D>(parameters, "point1", new Point3D(0, 0, 0));
            var point2 = GetParameter<Point3D>(parameters, "point2", new Point3D(5, 0, 0));
            var dimLine = GetParameter<Point3D>(parameters, "dimLine", new Point3D(2.5, 2, 0));

            autoCADCode.AppendLine($"                AlignedDimension dim = new AlignedDimension();");
            autoCADCode.AppendLine($"                dim.XLine1Point = new Point3d({point1.X}, {point1.Y}, {point1.Z});");
            autoCADCode.AppendLine($"                dim.XLine2Point = new Point3d({point2.X}, {point2.Y}, {point2.Z});");
            autoCADCode.AppendLine($"                dim.DimLinePoint = new Point3d({dimLine.X}, {dimLine.Y}, {dimLine.Z});");
            autoCADCode.AppendLine($"                btr.AppendEntity(dim);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(dim, true);");
        }

        private void GenerateAngularDimensionCode(Dictionary<string, object> parameters)
        {
            var center = GetParameter<Point3D>(parameters, "center", new Point3D(0, 0, 0));
            var point1 = GetParameter<Point3D>(parameters, "point1", new Point3D(5, 0, 0));
            var point2 = GetParameter<Point3D>(parameters, "point2", new Point3D(0, 5, 0));
            var arcPoint = GetParameter<Point3D>(parameters, "arcPoint", new Point3D(3, 3, 0));

            autoCADCode.AppendLine($"                Point3AngularDimension angDim = new Point3AngularDimension();");
            autoCADCode.AppendLine($"                angDim.CenterPoint = new Point3d({center.X}, {center.Y}, {center.Z});");
            autoCADCode.AppendLine($"                angDim.XLine1Point = new Point3d({point1.X}, {point1.Y}, {point1.Z});");
            autoCADCode.AppendLine($"                angDim.XLine2Point = new Point3d({point2.X}, {point2.Y}, {point2.Z});");
            autoCADCode.AppendLine($"                angDim.ArcPoint = new Point3d({arcPoint.X}, {arcPoint.Y}, {arcPoint.Z});");
            autoCADCode.AppendLine($"                btr.AppendEntity(angDim);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(angDim, true);");
        }

        private void GenerateRadialDimensionCode(Dictionary<string, object> parameters)
        {
            var center = GetParameter<Point3D>(parameters, "center", new Point3D(0, 0, 0));
            var chordPoint = GetParameter<Point3D>(parameters, "chordPoint", new Point3D(5, 0, 0));
            var leaderLength = GetParameter<double>(parameters, "leaderLength", 2.0);

            autoCADCode.AppendLine($"                RadialDimension radDim = new RadialDimension();");
            autoCADCode.AppendLine($"                radDim.Center = new Point3d({center.X}, {center.Y}, {center.Z});");
            autoCADCode.AppendLine($"                radDim.ChordPoint = new Point3d({chordPoint.X}, {chordPoint.Y}, {chordPoint.Z});");
            autoCADCode.AppendLine($"                radDim.LeaderLength = {leaderLength};");
            autoCADCode.AppendLine($"                btr.AppendEntity(radDim);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(radDim, true);");
        }

        private void GenerateDiametricDimensionCode(Dictionary<string, object> parameters)
        {
            var chordPoint = GetParameter<Point3D>(parameters, "chordPoint", new Point3D(5, 0, 0));
            var farChordPoint = GetParameter<Point3D>(parameters, "farChordPoint", new Point3D(-5, 0, 0));
            var leaderLength = GetParameter<double>(parameters, "leaderLength", 2.0);

            autoCADCode.AppendLine($"                DiametricDimension diamDim = new DiametricDimension();");
            autoCADCode.AppendLine($"                diamDim.ChordPoint = new Point3d({chordPoint.X}, {chordPoint.Y}, {chordPoint.Z});");
            autoCADCode.AppendLine($"                diamDim.FarChordPoint = new Point3d({farChordPoint.X}, {farChordPoint.Y}, {farChordPoint.Z});");
            autoCADCode.AppendLine($"                diamDim.LeaderLength = {leaderLength};");
            autoCADCode.AppendLine($"                btr.AppendEntity(diamDim);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(diamDim, true);");
        }

        // Surfaces and Regions implementation
        private void GenerateHatchCode(Dictionary<string, object> parameters)
        {
            var pattern = GetParameter<string>(parameters, "pattern", "SOLID");
            var scale = GetParameter<double>(parameters, "scale", 1.0);
            var angle = GetParameter<double>(parameters, "angle", 0.0);

            autoCADCode.AppendLine($"                Hatch hatch = new Hatch();");
            autoCADCode.AppendLine($"                hatch.SetHatchPattern(HatchPatternType.PreDefined, \"{pattern}\");");
            autoCADCode.AppendLine($"                hatch.PatternScale = {scale};");
            autoCADCode.AppendLine($"                hatch.PatternAngle = {angle * Math.PI / 180.0};");
            autoCADCode.AppendLine($"                btr.AppendEntity(hatch);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(hatch, true);");
            autoCADCode.AppendLine($"                // Note: Boundary must be added after creation");
        }

        private void GenerateRegionCode(Dictionary<string, object> parameters)
        {
            autoCADCode.AppendLine($"                // Region creation requires existing curves");
            autoCADCode.AppendLine($"                // Use Region.CreateFromCurves() with curve collection");
        }

        private void GenerateSurfaceCode(Dictionary<string, object> parameters)
        {
            var width = GetParameter<double>(parameters, "width", 10.0);
            var height = GetParameter<double>(parameters, "height", 10.0);
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                PlaneSurface surface = new PlaneSurface();");
            autoCADCode.AppendLine($"                surface.CreateByWidthHeight(new Point3d({position.X}, {position.Y}, {position.Z}), ");
            autoCADCode.AppendLine($"                    Vector3d.XAxis, Vector3d.YAxis, {width}, {height});");
            autoCADCode.AppendLine($"                btr.AppendEntity(surface);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(surface, true);");
        }

        private void GenerateMeshCode(Dictionary<string, object> parameters)
        {
            var subdivisionLevel = GetParameter<double>(parameters, "subdivisionLevel", 0);

            autoCADCode.AppendLine($"                SubDMesh mesh = new SubDMesh();");
            autoCADCode.AppendLine($"                mesh.SubDivisionLevel = (int){subdivisionLevel};");
            autoCADCode.AppendLine($"                btr.AppendEntity(mesh);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(mesh, true);");
            autoCADCode.AppendLine($"                // Note: Vertices and faces must be added after creation");
        }

        // Points and Blocks implementation
        private void GeneratePointCode(Dictionary<string, object> parameters)
        {
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                DBPoint point = new DBPoint();");
            autoCADCode.AppendLine($"                point.Position = new Point3d({position.X}, {position.Y}, {position.Z});");
            autoCADCode.AppendLine($"                btr.AppendEntity(point);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(point, true);");
        }

        private void GenerateBlockRefCode(Dictionary<string, object> parameters)
        {
            var blockName = GetParameter<string>(parameters, "blockName", "TestBlock");
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));
            var scale = GetParameter<double>(parameters, "scale", 1.0);
            var rotation = GetParameter<double>(parameters, "rotation", 0.0) * Math.PI / 180.0;

            autoCADCode.AppendLine($"                // Block reference requires existing block definition");
            autoCADCode.AppendLine($"                BlockTable blockTable = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);");
            autoCADCode.AppendLine($"                if (blockTable.Has(\"{blockName}\"))");
            autoCADCode.AppendLine($"                {{");
            autoCADCode.AppendLine($"                    BlockReference blockRef = new BlockReference(");
            autoCADCode.AppendLine($"                        new Point3d({position.X}, {position.Y}, {position.Z}), blockTable[\"{blockName}\"]);");
            autoCADCode.AppendLine($"                    blockRef.ScaleFactors = new Scale3d({scale});");
            autoCADCode.AppendLine($"                    blockRef.Rotation = {rotation};");
            autoCADCode.AppendLine($"                    btr.AppendEntity(blockRef);");
            autoCADCode.AppendLine($"                    tr.AddNewlyCreatedDBObject(blockRef, true);");
            autoCADCode.AppendLine($"                }}");
        }

        // Organization implementation
        private void GenerateLayerCode(Dictionary<string, object> parameters)
        {
            var name = GetParameter<string>(parameters, "name", "Layer1");
            var color = GetParameter<Color3D>(parameters, "color", new Color3D(255, 255, 255));

            autoCADCode.AppendLine($"                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForWrite);");
            autoCADCode.AppendLine($"                if (!lt.Has(\"{name}\"))");
            autoCADCode.AppendLine($"                {{");
            autoCADCode.AppendLine($"                    LayerTableRecord ltr = new LayerTableRecord();");
            autoCADCode.AppendLine($"                    ltr.Name = \"{name}\";");
            autoCADCode.AppendLine($"                    ltr.Color = Color.FromRgb({color.R}, {color.G}, {color.B});");
            autoCADCode.AppendLine($"                    lt.Add(ltr);");
            autoCADCode.AppendLine($"                    tr.AddNewlyCreatedDBObject(ltr, true);");
            autoCADCode.AppendLine($"                }}");
        }

        private void GenerateGroupCode(Dictionary<string, object> parameters)
        {
            var name = GetParameter<string>(parameters, "name", "Group1");
            var description = GetParameter<string>(parameters, "description", "Generated Group");

            autoCADCode.AppendLine($"                Group group = new Group(\"{description}\", true);");
            autoCADCode.AppendLine($"                DBDictionary groupDict = (DBDictionary)tr.GetObject(db.GroupDictionaryId, OpenMode.ForWrite);");
            autoCADCode.AppendLine($"                groupDict.SetAt(\"{name}\", group);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(group, true);");
        }

        // Transformations implementation
        private void GenerateMoveCode(Dictionary<string, object> parameters)
        {
            var from = GetParameter<Point3D>(parameters, "from", new Point3D(0, 0, 0));
            var to = GetParameter<Point3D>(parameters, "to", new Point3D(1, 1, 0));

            autoCADCode.AppendLine($"                Vector3d moveVector = new Vector3d({to.X - from.X}, {to.Y - from.Y}, {to.Z - from.Z});");
            autoCADCode.AppendLine($"                Matrix3d moveMatrix = Matrix3d.Displacement(moveVector);");
            autoCADCode.AppendLine($"                // Apply to selected objects: entity.TransformBy(moveMatrix);");
        }

        private void GenerateRotateCode(Dictionary<string, object> parameters)
        {
            var center = GetParameter<Point3D>(parameters, "center", new Point3D(0, 0, 0));
            var angle = GetParameter<double>(parameters, "angle", 0.0) * Math.PI / 180.0;
            var axis = GetParameter<Vector3D>(parameters, "axis", new Vector3D(0, 0, 1));

            autoCADCode.AppendLine($"                Point3d rotationCenter = new Point3d({center.X}, {center.Y}, {center.Z});");
            autoCADCode.AppendLine($"                Vector3d rotationAxis = new Vector3d({axis.X}, {axis.Y}, {axis.Z});");
            autoCADCode.AppendLine($"                Matrix3d rotationMatrix = Matrix3d.Rotation({angle}, rotationAxis, rotationCenter);");
            autoCADCode.AppendLine($"                // Apply to selected objects: entity.TransformBy(rotationMatrix);");
        }

        private void GenerateScaleCode(Dictionary<string, object> parameters)
        {
            var center = GetParameter<Point3D>(parameters, "center", new Point3D(0, 0, 0));
            var factor = GetParameter<double>(parameters, "factor", 1.0);

            autoCADCode.AppendLine($"                Point3d scaleCenter = new Point3d({center.X}, {center.Y}, {center.Z});");
            autoCADCode.AppendLine($"                Matrix3d scaleMatrix = Matrix3d.Scaling({factor}, scaleCenter);");
            autoCADCode.AppendLine($"                // Apply to selected objects: entity.TransformBy(scaleMatrix);");
        }

        private void GenerateMirrorCode(Dictionary<string, object> parameters)
        {
            var point1 = GetParameter<Point3D>(parameters, "point1", new Point3D(0, 0, 0));
            var point2 = GetParameter<Point3D>(parameters, "point2", new Point3D(1, 0, 0));

            autoCADCode.AppendLine($"                Point3d mirrorPt1 = new Point3d({point1.X}, {point1.Y}, {point1.Z});");
            autoCADCode.AppendLine($"                Point3d mirrorPt2 = new Point3d({point2.X}, {point2.Y}, {point2.Z});");
            autoCADCode.AppendLine($"                Line3d mirrorLine = new Line3d(mirrorPt1, mirrorPt2);");
            autoCADCode.AppendLine($"                Matrix3d mirrorMatrix = Matrix3d.Mirroring(mirrorLine);");
            autoCADCode.AppendLine($"                // Apply to selected objects: entity.TransformBy(mirrorMatrix);");
        }

        private void GenerateArrayRectangularCode(Dictionary<string, object> parameters)
        {
            var rows = GetParameter<double>(parameters, "rows", 3);
            var cols = GetParameter<double>(parameters, "cols", 3);
            var rowSpacing = GetParameter<double>(parameters, "rowSpacing", 5.0);
            var colSpacing = GetParameter<double>(parameters, "colSpacing", 5.0);

            autoCADCode.AppendLine($"                // Rectangular array - repeat for each row and column");
            autoCADCode.AppendLine($"                for (int row = 0; row < {rows}; row++)");
            autoCADCode.AppendLine($"                {{");
            autoCADCode.AppendLine($"                    for (int col = 0; col < {cols}; col++)");
            autoCADCode.AppendLine($"                    {{");
            autoCADCode.AppendLine($"                        Vector3d offset = new Vector3d(col * {colSpacing}, row * {rowSpacing}, 0);");
            autoCADCode.AppendLine($"                        Matrix3d transform = Matrix3d.Displacement(offset);");
            autoCADCode.AppendLine($"                        // Clone and transform entity");
            autoCADCode.AppendLine($"                    }}");
            autoCADCode.AppendLine($"                }}");
        }

        private void GenerateArrayPolarCode(Dictionary<string, object> parameters)
        {
            var center = GetParameter<Point3D>(parameters, "center", new Point3D(0, 0, 0));
            var count = GetParameter<double>(parameters, "count", 6);
            var angle = GetParameter<double>(parameters, "angle", 360.0) * Math.PI / 180.0;

            autoCADCode.AppendLine($"                Point3d arrayCenter = new Point3d({center.X}, {center.Y}, {center.Z});");
            autoCADCode.AppendLine($"                double angleStep = {angle} / {count};");
            autoCADCode.AppendLine($"                for (int i = 0; i < {count}; i++)");
            autoCADCode.AppendLine($"                {{");
            autoCADCode.AppendLine($"                    double currentAngle = i * angleStep;");
            autoCADCode.AppendLine($"                    Matrix3d transform = Matrix3d.Rotation(currentAngle, Vector3d.ZAxis, arrayCenter);");
            autoCADCode.AppendLine($"                    // Clone and transform entity");
            autoCADCode.AppendLine($"                }}");
        }

        // Boolean Operations implementation
        private void GenerateUnionCode(Dictionary<string, object> parameters)
        {
            autoCADCode.AppendLine($"                // Boolean Union - requires two Solid3d objects");
            autoCADCode.AppendLine($"                // solid1.BooleanOperation(BooleanOperationType.BoolUnite, solid2);");
        }

        private void GenerateSubtractCode(Dictionary<string, object> parameters)
        {
            autoCADCode.AppendLine($"                // Boolean Subtract - requires two Solid3d objects");
            autoCADCode.AppendLine($"                // solid1.BooleanOperation(BooleanOperationType.BoolSubtract, solid2);");
        }

        private void GenerateIntersectCode(Dictionary<string, object> parameters)
        {
            autoCADCode.AppendLine($"                // Boolean Intersect - requires two Solid3d objects");
            autoCADCode.AppendLine($"                // solid1.BooleanOperation(BooleanOperationType.BoolIntersect, solid2);");
        }

        // Materials and Visualization implementation
        private void GenerateMaterialCode(Dictionary<string, object> parameters)
        {
            var name = GetParameter<string>(parameters, "name", "CustomMaterial");
            var color = GetParameter<Color3D>(parameters, "color", new Color3D(128, 128, 128));

            autoCADCode.AppendLine($"                // Material creation");
            autoCADCode.AppendLine($"                Material material = new Material();");
            autoCADCode.AppendLine($"                material.Name = \"{name}\";");
            autoCADCode.AppendLine($"                material.Diffuse = new MaterialColor(Color.FromRgb({color.R}, {color.G}, {color.B}));");
        }

        private void GenerateLightCode(Dictionary<string, object> parameters)
        {
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 10));
            var intensity = GetParameter<double>(parameters, "intensity", 1.0);

            autoCADCode.AppendLine($"                Light light = new Light();");
            autoCADCode.AppendLine($"                light.Position = new Point3d({position.X}, {position.Y}, {position.Z});");
            autoCADCode.AppendLine($"                light.Intensity = {intensity};");
            autoCADCode.AppendLine($"                btr.AppendEntity(light);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(light, true);");
        }

        private void GenerateCameraCode(Dictionary<string, object> parameters)
        {
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 10));
            var target = GetParameter<Point3D>(parameters, "target", new Point3D(0, 0, 0));

            autoCADCode.AppendLine($"                // Camera setup");
            autoCADCode.AppendLine($"                ViewTableRecord view = new ViewTableRecord();");
            autoCADCode.AppendLine($"                view.CenterPoint = new Point2d({target.X}, {target.Y});");
            autoCADCode.AppendLine($"                view.Target = new Point3d({target.X}, {target.Y}, {target.Z});");
        }

        // Advanced Entities implementation
        private void GenerateLeaderCode(Dictionary<string, object> parameters)
        {
            var start = GetParameter<Point3D>(parameters, "start", new Point3D(0, 0, 0));
            var end = GetParameter<Point3D>(parameters, "end", new Point3D(5, 5, 0));

            autoCADCode.AppendLine($"                Leader leader = new Leader();");
            autoCADCode.AppendLine($"                leader.AppendVertex(new Point3d({start.X}, {start.Y}, {start.Z}));");
            autoCADCode.AppendLine($"                leader.AppendVertex(new Point3d({end.X}, {end.Y}, {end.Z}));");
            autoCADCode.AppendLine($"                btr.AppendEntity(leader);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(leader, true);");
        }

        private void GenerateMLeaderCode(Dictionary<string, object> parameters)
        {
            var position = GetParameter<Point3D>(parameters, "position", new Point3D(0, 0, 0));
            var text = GetParameter<string>(parameters, "text", "Leader Text");

            autoCADCode.AppendLine($"                MLeader mleader = new MLeader();");
            autoCADCode.AppendLine($"                mleader.SetDatabaseDefaults();");
            autoCADCode.AppendLine($"                // MLeader setup requires more complex configuration");
            autoCADCode.AppendLine($"                btr.AppendEntity(mleader);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(mleader, true);");
        }

        private void GenerateMlineCode(Dictionary<string, object> parameters)
        {
            var start = GetParameter<Point3D>(parameters, "start", new Point3D(0, 0, 0));
            var end = GetParameter<Point3D>(parameters, "end", new Point3D(10, 0, 0));

            autoCADCode.AppendLine($"                Mline mline = new Mline();");
            autoCADCode.AppendLine($"                mline.AppendVertex(new Point3d({start.X}, {start.Y}, {start.Z}));");
            autoCADCode.AppendLine($"                mline.AppendVertex(new Point3d({end.X}, {end.Y}, {end.Z}));");
            autoCADCode.AppendLine($"                btr.AppendEntity(mline);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(mline, true);");
        }

        private void GenerateTraceCode(Dictionary<string, object> parameters)
        {
            var point1 = GetParameter<Point3D>(parameters, "point1", new Point3D(0, 0, 0));
            var point2 = GetParameter<Point3D>(parameters, "point2", new Point3D(1, 0, 0));
            var point3 = GetParameter<Point3D>(parameters, "point3", new Point3D(1, 1, 0));
            var point4 = GetParameter<Point3D>(parameters, "point4", new Point3D(0, 1, 0));

            autoCADCode.AppendLine($"                Trace trace = new Trace();");
            autoCADCode.AppendLine($"                trace.SetPointAt(0, new Point3d({point1.X}, {point1.Y}, {point1.Z}));");
            autoCADCode.AppendLine($"                trace.SetPointAt(1, new Point3d({point2.X}, {point2.Y}, {point2.Z}));");
            autoCADCode.AppendLine($"                trace.SetPointAt(2, new Point3d({point3.X}, {point3.Y}, {point3.Z}));");
            autoCADCode.AppendLine($"                trace.SetPointAt(3, new Point3d({point4.X}, {point4.Y}, {point4.Z}));");
            autoCADCode.AppendLine($"                btr.AppendEntity(trace);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(trace, true);");
        }

        private void GenerateSolid2DCode(Dictionary<string, object> parameters)
        {
            var point1 = GetParameter<Point3D>(parameters, "point1", new Point3D(0, 0, 0));
            var point2 = GetParameter<Point3D>(parameters, "point2", new Point3D(1, 0, 0));
            var point3 = GetParameter<Point3D>(parameters, "point3", new Point3D(1, 1, 0));
            var point4 = GetParameter<Point3D>(parameters, "point4", new Point3D(0, 1, 0));

            autoCADCode.AppendLine($"                Solid solid2d = new Solid();");
            autoCADCode.AppendLine($"                solid2d.SetPointAt(0, new Point3d({point1.X}, {point1.Y}, {point1.Z}));");
            autoCADCode.AppendLine($"                solid2d.SetPointAt(1, new Point3d({point2.X}, {point2.Y}, {point2.Z}));");
            autoCADCode.AppendLine($"                solid2d.SetPointAt(2, new Point3d({point4.X}, {point4.Y}, {point4.Z})); // Note: order is important");");
            autoCADCode.AppendLine($"                solid2d.SetPointAt(3, new Point3d({point3.X}, {point3.Y}, {point3.Z}));");
            autoCADCode.AppendLine($"                btr.AppendEntity(solid2d);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(solid2d, true);");
        }

        private void GenerateFaceCode(Dictionary<string, object> parameters)
        {
            var point1 = GetParameter<Point3D>(parameters, "point1", new Point3D(0, 0, 0));
            var point2 = GetParameter<Point3D>(parameters, "point2", new Point3D(1, 0, 0));
            var point3 = GetParameter<Point3D>(parameters, "point3", new Point3D(1, 1, 0));
            var point4 = GetParameter<Point3D>(parameters, "point4", new Point3D(0, 1, 0));

            autoCADCode.AppendLine($"                Face face = new Face();");
            autoCADCode.AppendLine($"                face.SetVertexAt(0, new Point3d({point1.X}, {point1.Y}, {point1.Z}));");
            autoCADCode.AppendLine($"                face.SetVertexAt(1, new Point3d({point2.X}, {point2.Y}, {point2.Z}));");
            autoCADCode.AppendLine($"                face.SetVertexAt(2, new Point3d({point3.X}, {point3.Y}, {point3.Z}));");
            autoCADCode.AppendLine($"                face.SetVertexAt(3, new Point3d({point4.X}, {point4.Y}, {point4.Z}));");
            autoCADCode.AppendLine($"                btr.AppendEntity(face);");
            autoCADCode.AppendLine($"                tr.AddNewlyCreatedDBObject(face, true);");
        }

        private T GetParameter<T>(Dictionary<string, object> parameters, string key, T defaultValue)
        {
            if (parameters.ContainsKey(key))
            {
                var value = parameters[key];
                if (value is T directValue)
                    return directValue;
                
                // Try conversion for compatible types
                try
                {
                    return (T)Convert.ChangeType(value, typeof(T));
                }
                catch
                {
                    return defaultValue;
                }
            }
            return defaultValue;
        }

        // File operations
        public void SaveScript(string filename, string script)
        {
            File.WriteAllText(filename, script);
            Console.WriteLine($"Script saved to {filename}");
        }

        public string LoadScript(string filename)
        {
            if (!File.Exists(filename))
                throw new FileNotFoundException($"Script file not found: {filename}");
            
            string script = File.ReadAllText(filename);
            Console.WriteLine($"Script loaded from {filename}");
            return script;
        }

        public void SaveAutoCADCode(string filename)
        {
            string code = GenerateAutoCADCode();
            File.WriteAllText(filename, code);
            Console.WriteLine($"AutoCAD code saved to {filename}");
        }

        public void Clear()
        {
            commands.Clear();
            variables.Clear();
            autoCADCode.Clear();
        }

        public void ListCommands()
        {
            Console.WriteLine($"Total commands: {commands.Count}");
            for (int i = 0; i < commands.Count; i++)
            {
                var cmd = commands[i];
                Console.WriteLine($"{i + 1}. Line {cmd.LineNumber}: {cmd.Type} ({cmd.Parameters.Count} params)");
            }
        }

        public void ListVariables()
        {
            Console.WriteLine($"Variables ({variables.Count}):");
            foreach (var var in variables)
            {
                Console.WriteLine($"  ${var.Key} = {var.Value}");
            }
        }

        public void ShowHelp()
        {
            Console.WriteLine("\n=== CAD INTERPRETER COMPLETE COMMAND REFERENCE ===\n");

            Console.WriteLine("3D SOLIDS:");
            Console.WriteLine("  BOX(width=10, height=5, depth=3, position=[0,0,0])");
            Console.WriteLine("  SPHERE(radius=2.5, position=[0,0,0])");
            Console.WriteLine("  CYLINDER(radius=1, height=8, position=[0,0,0])");
            Console.WriteLine("  CONE(radius=2, height=6, position=[0,0,0])");
            Console.WriteLine("  TORUS(majorRadius=2, minorRadius=0.5, position=[0,0,0])");
            Console.WriteLine("  WEDGE(length=2, width=1, height=1, position=[0,0,0])");
            Console.WriteLine("  PYRAMID(radius=1, height=2, sides=4, position=[0,0,0])");

            Console.WriteLine("\nCURVES & LINES:");
            Console.WriteLine("  LINE(start=[0,0,0], end=[10,0,0])");
            Console.WriteLine("  CIRCLE(radius=5, center=[0,0,0], normal=<0,0,1>)");
            Console.WriteLine("  ARC(radius=3, center=[0,0,0], startAngle=0, endAngle=90, normal=<0,0,1>)");
            Console.WriteLine("  ELLIPSE(center=[0,0,0], majorAxis=<2,0,0>, radiusRatio=0.5, normal=<0,0,1>)");
            Console.WriteLine("  SPLINE(start=[0,0,0], end=[5,5,0], control1=[1,2,0], control2=[4,3,0])");
            Console.WriteLine("  POLYLINE(start=[0,0,0], end=[5,5,0], width=0.5)");
            Console.WriteLine("  POLYLINE3D(start=[0,0,0], end=[5,5,5])");
            Console.WriteLine("  RAY(basePoint=[0,0,0], direction=<1,1,0>)");
            Console.WriteLine("  XLINE(basePoint=[0,0,0], direction=<1,0,0>)");

            Console.WriteLine("\nTEXT & ANNOTATIONS:");
            Console.WriteLine("  TEXT(text=\"Hello\", position=[0,0,0], height=1, rotation=0)");
            Console.WriteLine("  MTEXT(text=\"Multi-line text\", position=[0,0,0], height=1, width=10)");

            Console.WriteLine("\nDIMENSIONS:");
            Console.WriteLine("  DIMENSION_LINEAR(point1=[0,0,0], point2=[5,0,0], dimLine=[2.5,2,0])");
            Console.WriteLine("  DIMENSION_ANGULAR(center=[0,0,0], point1=[5,0,0], point2=[0,5,0], arcPoint=[3,3,0])");
            Console.WriteLine("  DIMENSION_RADIAL(center=[0,0,0], chordPoint=[5,0,0], leaderLength=2)");
            Console.WriteLine("  DIMENSION_DIAMETRIC(chordPoint=[5,0,0], farChordPoint=[-5,0,0], leaderLength=2)");

            Console.WriteLine("\nSURFACES & REGIONS:");
            Console.WriteLine("  HATCH(pattern=\"SOLID\", scale=1, angle=0)");
            Console.WriteLine("  REGION() // Creates region from existing curves");
            Console.WriteLine("  SURFACE(width=10, height=10, position=[0,0,0])");
            Console.WriteLine("  MESH(subdivisionLevel=0)");

            Console.WriteLine("\nPOINTS & BLOCKS:");
            Console.WriteLine("  POINT(position=[0,0,0])");
            Console.WriteLine("  BLOCK_REF(blockName=\"TestBlock\", position=[0,0,0], scale=1, rotation=0)");

            Console.WriteLine("\nORGANIZATION:");
            Console.WriteLine("  LAYER(name=\"Layer1\", color=RGB(255,0,0))");
            Console.WriteLine("  GROUP(name=\"Group1\", description=\"My Group\")");

            Console.WriteLine("\nTRANSFORMATIONS:");
            Console.WriteLine("  MOVE(from=[0,0,0], to=[10,5,0])");
            Console.WriteLine("  ROTATE(center=[0,0,0], angle=45, axis=<0,0,1>)");
            Console.WriteLine("  SCALE(center=[0,0,0], factor=2.0)");
            Console.WriteLine("  MIRROR(point1=[0,0,0], point2=[1,0,0])");
            Console.WriteLine("  ARRAY_RECTANGULAR(rows=3, cols=4, rowSpacing=5, colSpacing=5)");
            Console.WriteLine("  ARRAY_POLAR(center=[0,0,0], count=6, angle=360)");

            Console.WriteLine("\nBOOLEAN OPERATIONS:");
            Console.WriteLine("  UNION() // Combines two solids");
            Console.WriteLine("  SUBTRACT() // Subtracts one solid from another");
            Console.WriteLine("  INTERSECT() // Creates intersection of two solids");

            Console.WriteLine("\nMATERIALS & VISUALIZATION:");
            Console.WriteLine("  MATERIAL(name=\"Steel\", color=RGB(128,128,128))");
            Console.WriteLine("  LIGHT(position=[0,0,10], intensity=1.0)");
            Console.WriteLine("  CAMERA(position=[0,0,10], target=[0,0,0])");

            Console.WriteLine("\nADVANCED ENTITIES:");
            Console.WriteLine("  LEADER(start=[0,0,0], end=[5,5,0])");
            Console.WriteLine("  MLEADER(position=[0,0,0], text=\"Leader Text\")");
            Console.WriteLine("  MLINE(start=[0,0,0], end=[10,0,0])");
            Console.WriteLine("  TRACE(point1=[0,0,0], point2=[1,0,0], point3=[1,1,0], point4=[0,1,0])");
            Console.WriteLine("  SOLID(point1=[0,0,0], point2=[1,0,0], point3=[1,1,0], point4=[0,1,0])");
            Console.WriteLine("  FACE(point1=[0,0,0], point2=[1,0,0], point3=[1,1,0], point4=[0,1,0])");

            Console.WriteLine("\nVARIABLES:");
            Console.WriteLine("  VAR name = value");
            Console.WriteLine("  Use variables with $name syntax");

            Console.WriteLine("\nCONTROL COMMANDS:");
            Console.WriteLine("  help - Show this help");
            Console.WriteLine("  execute - Process all commands");
            Console.WriteLine("  list - Show parsed commands");
            Console.WriteLine("  vars - Show variables");
            Console.WriteLine("  clear - Clear all commands and variables");
            Console.WriteLine("  generate - Show generated AutoCAD code");
            Console.WriteLine("  save filename.cs - Save AutoCAD code to file");
            Console.WriteLine("  load filename.txt - Load and parse script file");
            Console.WriteLine("  quit - Exit interpreter");

            Console.WriteLine("\nDATA TYPES:");
            Console.WriteLine("  Points: [x,y,z]");
            Console.WriteLine("  Vectors: <x,y,z>");
            Console.WriteLine("  Colors: RGB(r,g,b)");
            Console.WriteLine("  Strings: \"text\"");
            Console.WriteLine("  Numbers: 1.5, 42");
            Console.WriteLine("  Booleans: true, false");
        }
    }

    // Enhanced test program with complete functionality
    class Program
    {
        static void Main(string[] args)
        {
            var interpreter = new CADInterpreter();

            Console.WriteLine("===== COMPLETE CAD COMMAND INTERPRETER =====");
            Console.WriteLine("Version 2.0 - Full AutoCAD .NET API Support");
            Console.WriteLine("============================================");
            Console.WriteLine();
            Console.WriteLine("Type 'help' for complete command reference");
            Console.WriteLine("Type 'quit' to exit");
            Console.WriteLine();

            // Demo script
            Console.WriteLine("Demo Script - Building with Materials:");
            Console.WriteLine("--------------------------------------");
            var demoScript = @"
VAR wallHeight = 3
VAR wallLength = 10
VAR wallThickness = 0.3

// Create foundation
BOX(width=$wallLength, height=$wallLength, depth=0.5, position=[0,0,0])

// Create walls
BOX(width=$wallLength, height=$wallThickness, depth=$wallHeight, position=[0,0,0.5])
BOX(width=$wallLength, height=$wallThickness, depth=$wallHeight, position=[0,$wallLength,0.5])
BOX(width=$wallThickness, height=$wallLength, depth=$wallHeight, position=[0,0,0.5])
BOX(width=$wallThickness, height=$wallLength, depth=$wallHeight, position=[$wallLength,0,0.5])

// Add roof structure
PYRAMID(radius=7, height=2, sides=4, position=[5,5,3.5])

// Create layers and materials
LAYER(name=""structure"", color=RGB(128,128,128))
LAYER(name=""openings"", color=RGB(255,255,0))
MATERIAL(name=""concrete"", color=RGB(200,200,200))

// Add annotations
TEXT(text=""Building Plan"", position=[5,12,0], height=0.8)
DIMENSION_LINEAR(point1=[0,0,0], point2=[10,0,0], dimLine=[5,-2,0])

// Add circular elements
CIRCLE(radius=1, center=[5,5,0])
TORUS(majorRadius=1.5, minorRadius=0.3, position=[2,2,4])
";

            Console.WriteLine(demoScript.Trim());
            Console.WriteLine();

            while (true)
            {
                Console.Write("> ");
                var input = Console.ReadLine();

                try
                {
                    if (input?.ToLower() == "quit")
                        break;
                    else if (input?.ToLower() == "help")
                        interpreter.ShowHelp();
                    else if (input?.ToLower() == "demo")
                    {
                        interpreter.ParseScript(demoScript);
                        Console.WriteLine("Demo script loaded!");
                    }
                    else if (input?.ToLower() == "execute")
                        Console.WriteLine("Commands parsed successfully!");
                    else if (input?.ToLower() == "list")
                        interpreter.ListCommands();
                    else if (input?.ToLower() == "vars")
                        interpreter.ListVariables();
                    else if (input?.ToLower() == "clear")
                        interpreter.Clear();
                    else if (input?.ToLower().StartsWith("save ") == true)
                    {
                        var filename = input.Substring(5).Trim();
                        interpreter.SaveAutoCADCode(filename);
                    }
                    else if (input?.ToLower().StartsWith("load ") == true)
                    {
                        var filename = input.Substring(5).Trim();
                        var script = interpreter.LoadScript(filename);
                        interpreter.ParseScript(script);
                    }
                    else if (input?.ToLower() == "generate")
                    {
                        var code = interpreter.GenerateAutoCADCode();
                        Console.WriteLine("Generated AutoCAD Code:");
                        Console.WriteLine("========================");
                        Console.WriteLine(code);
                    }
                    else if (!string.IsNullOrEmpty(input))
                        interpreter.ParseCommand(input);
                }
                catch (InterpreterException ex)
                {
                    Console.WriteLine($"Error: {ex.Message}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error: {ex.Message}");
                }
            }

            Console.WriteLine("Goodbye!");
        }
    }
}
