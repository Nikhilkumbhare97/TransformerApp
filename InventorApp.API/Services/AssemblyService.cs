using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Inventor;
using InventorApp.API.Models;
using System.Windows.Forms;

namespace InventorApp.API.Services
{
    public class AssemblyService
    {
        private Inventor.Application? _inventorApp;
        private bool _isAssemblyOpen = false;

        public bool IsAssemblyOpen => _isAssemblyOpen;

        private Inventor.Application GetInventorApplication()
        {
            if (_inventorApp == null)
            {
                try
                {
                    Type? inventorType = Type.GetTypeFromProgID("Inventor.Application");
                    if (inventorType == null)
                    {
                        throw new InvalidOperationException("Autodesk Inventor is not installed or registered. Please ensure Inventor is installed and properly registered.");
                    }

                    _inventorApp = (Inventor.Application)Activator.CreateInstance(inventorType)!;
                    _inventorApp.Visible = true;
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to initialize Inventor application: {ex.Message}. Please ensure Inventor is running and properly registered.", ex);
                }
            }
            return _inventorApp;
        }

        public void OpenAssembly(string assemblyPath)
        {
            try
            {
                if (_inventorApp == null)
                {
                    Type? inventorType = Type.GetTypeFromProgID("Inventor.Application");
                    if (inventorType == null) throw new InvalidOperationException("Autodesk Inventor is not installed or registered.");

                    _inventorApp = (Inventor.Application)Activator.CreateInstance(inventorType)!;
                    _inventorApp.Visible = true;
                }

                Documents docs = _inventorApp.Documents;
                docs.Open(assemblyPath);
                _isAssemblyOpen = true;
                Console.WriteLine($"Opening assembly: {assemblyPath}");
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error opening assembly: {e.Message}");
                throw;
            }
        }

        public void CloseAssembly()
        {
            try
            {
                if (_isAssemblyOpen && _inventorApp != null)
                {
                    _inventorApp.ActiveDocument.Close(false);
                    _isAssemblyOpen = false;
                    Console.WriteLine("Closing assembly...");
                }
                else
                {
                    Console.Error.WriteLine("No active application instance to close.");
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error closing assembly: {e.Message}");
                throw;
            }
            finally
            {
                if (_inventorApp != null)
                {
                    _inventorApp.Quit();
                    Marshal.ReleaseComObject(_inventorApp);
                    _inventorApp = null;
                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
        }

        public void ChangeParameters(string partFilePath, List<Dictionary<string, object>> parameters)
        {
            Document? partDoc = null;
            try
            {
                var inventorApp = GetInventorApplication();
                inventorApp.SilentOperation = true; // Suppress dialogs

                partDoc = inventorApp.Documents.Open(partFilePath, true); // Open with full access
                PartDocument part = (PartDocument)partDoc;
                Parameters paramList = part.ComponentDefinition.Parameters;

                foreach (var param in parameters)
                {
                    if (param.TryGetValue("parameterName", out var paramNameObj) && paramNameObj != null && param.TryGetValue("newValue", out var newValueObj))
                    {
                        string paramName = paramNameObj.ToString()!;
                        if (double.TryParse(newValueObj.ToString(), out double newValue))
                        {
                            try
                            {
                                // First try to set the value directly without units
                                paramList[paramName].Expression = newValue.ToString();
                                Console.WriteLine($"Successfully set parameter '{paramName}' to {newValue}");
                            }
                            catch (Exception)
                            {
                                try
                                {
                                    // If that fails, try with units
                                    paramList[paramName].Expression = $"{newValue} mm";
                                    Console.WriteLine($"Successfully set parameter '{paramName}' to {newValue} mm");
                                }
                                catch (Exception unitEx)
                                {
                                    Console.Error.WriteLine($"Failed to set parameter '{paramName}' with value {newValue} (with and without units)");
                                    Console.Error.WriteLine($"Error details: {unitEx.Message}");
                                    throw new ArgumentException($"Failed to set parameter '{paramName}' with value {newValue}. Error: {unitEx.Message}", unitEx);
                                }
                            }
                        }
                        else
                        {
                            throw new ArgumentException($"Invalid value for parameter '{paramName}' in {partFilePath}.");
                        }
                    }
                    else
                    {
                        throw new ArgumentException("Missing 'parameterName' or 'newValue' in parameters.");
                    }
                }

                partDoc.Save2(true); // Save with Yes to All, suppress dialogs
                Console.WriteLine($"Parameters updated successfully in {partFilePath}");
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error changing parameters: {e.Message}");
                throw;
            }
            finally
            {
                if (_inventorApp != null)
                {
                    _inventorApp.SilentOperation = false; // Reset after operation
                }
                // Cleanup Inventor and COM objects
                try
                {
                    if (partDoc != null)
                    {
                        partDoc.Close(true); // Close and save
                        Marshal.ReleaseComObject(partDoc);
                    }

                    if (_inventorApp != null)
                    {
                        // Close all remaining documents
                        int maxAttempts = 2; // Prevent infinite loop
                        int currentAttempt = 0;
                        while (_inventorApp.Documents.Count > 0 && currentAttempt < maxAttempts)
                        {
                            try
                            {
                                Document doc = _inventorApp.Documents[1];
                                doc.Close(true);
                                Marshal.ReleaseComObject(doc);
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine($"Error closing document: {ex.Message}");
                                currentAttempt++;
                            }
                        }

                        if (currentAttempt >= maxAttempts)
                        {
                            Console.Error.WriteLine("Warning: Could not close all documents after maximum attempts");
                        }

                        // Quit Inventor and release COM object
                        _inventorApp.Quit();
                        Marshal.ReleaseComObject(_inventorApp);
                        _inventorApp = null;

                        // Force garbage collection to ensure COM objects are released
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error during cleanup: {ex.Message}");
                }
            }
        }

        public void UpdateIProperties(string filePath, Dictionary<string, string> properties)
        {
            try
            {
                if (_inventorApp == null)
                {
                    Type? inventorType = Type.GetTypeFromProgID("Inventor.Application");
                    if (inventorType == null) throw new InvalidOperationException("Autodesk Inventor is not installed or registered.");

                    _inventorApp = (Inventor.Application)Activator.CreateInstance(inventorType)!;
                    _inventorApp.Visible = false;
                }

                Document doc = _inventorApp.Documents.Open(filePath);
                PropertySets propSets = doc.PropertySets;

                foreach (var entry in properties)
                {
                    foreach (PropertySet set in propSets)
                    {
                        try
                        {
                            Property prop = set[entry.Key];
                            prop.Value = entry.Value;
                            break;
                        }
                        catch { }
                    }
                }

                doc.Save();
                doc.Close();
                Console.WriteLine($"Updated iProperties for {filePath}");
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error updating iProperties: {e.Message}");
                throw;
            }
        }

        public void UpdateIPropertiesForAssemblies(List<string> assemblyPaths, Dictionary<string, string> properties)
        {
            foreach (var path in assemblyPaths)
            {
                UpdateIProperties(path, properties);
            }
        }

        public void UpdateIpartsAndIassemblies(Dictionary<string, string> componentUpdates)
        {
            foreach (var kvp in componentUpdates)
            {
                SuppressComponent(kvp.Key, kvp.Value, false);
            }
        }

        public void SuppressMultipleComponents(List<SuppressAction> suppressActions)
        {
            try
            {
                var inventorApp = GetInventorApplication();

                foreach (var action in suppressActions)
                {
                    string assemblyPath = System.IO.Path.Combine("D:\\PROJECTS\\VECTOR\\3D Modelling\\TRANSFORMER\\WIP\\ABC099001\\MODEL", action.AssemblyFilePath);

                    foreach (var component in action.Components)
                    {
                        try
                        {
                            SuppressComponent(assemblyPath, component, action.Suppress);
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine($"Error suppressing component {component} in {assemblyPath}: {ex.Message}");
                            // Continue with next component
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error in SuppressMultipleComponents: {e.Message}");
                throw;
            }
            finally
            {
                // Cleanup Inventor and COM objects
                try
                {
                    if (_inventorApp != null)
                    {
                        // Close all open documents
                        while (_inventorApp.Documents.Count > 0)
                        {
                            try
                            {
                                Document doc = _inventorApp.Documents[1];
                                doc.Close(true);
                                Marshal.ReleaseComObject(doc);
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine($"Error closing document: {ex.Message}");
                            }
                        }

                        // Quit Inventor
                        _inventorApp.Quit();
                        Marshal.ReleaseComObject(_inventorApp);
                        _inventorApp = null;

                        // Force garbage collection to ensure COM objects are released
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error during cleanup: {ex.Message}");
                }
            }
        }

        public void SuppressComponent(string filePath, string componentName, bool suppress)
        {
            try
            {
                var inventorApp = GetInventorApplication();
                inventorApp.SilentOperation = true; // Suppress dialogs

                // Open the document
                Inventor.Document doc = inventorApp.Documents.Open(filePath, true);

                // Check the document type before casting
                if (doc is Inventor.AssemblyDocument asmDoc)
                {
                    SuppressAssemblyComponent(asmDoc, componentName, suppress);
                }
                else if (doc is Inventor.PartDocument partDoc)
                {
                    SuppressPartFeature(partDoc, componentName, suppress);
                }
                else
                {
                    throw new InvalidOperationException($"Unsupported document type: {doc.DocumentType}");
                }

                doc.Save2(true); // Save all changes, suppressing dialogs
                doc.Close(true); // Close and save
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error suppressing component: {e.Message}");
                throw;
            }
            finally
            {
                if (_inventorApp != null)
                    _inventorApp.SilentOperation = false; // Reset after operation
            }
        }

        private void SuppressAssemblyComponent(Inventor.AssemblyDocument asmDoc, string componentName, bool suppress)
        {
            ComponentOccurrences occurrences = asmDoc.ComponentDefinition.Occurrences;

            foreach (ComponentOccurrence occurrence in occurrences)
            {
                if (occurrence.Name.Equals(componentName, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Found assembly component: {occurrence.Name}");

                    if (suppress)
                        occurrence.Suppress();
                    else
                        occurrence.Unsuppress();

                    return; // Exit after suppressing the component
                }
            }

            throw new Exception($"Component '{componentName}' not found in assembly.");
        }

        private void SuppressPartFeature(Inventor.PartDocument partDoc, string featureName, bool suppress)
        {
            PartComponentDefinition partDef = partDoc.ComponentDefinition;

            // Find the feature in the part file
            foreach (PartFeature feature in partDef.Features)
            {
                if (feature.Name.Equals(featureName, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"Found part feature: {feature.Name}");

                    feature.Suppressed = suppress;
                    return;
                }
            }

            throw new Exception($"Feature '{featureName}' not found in part.");
        }

        public bool UpdateIPropertiesForAllFiles(string directoryPath, Dictionary<string, string> iProperties)
        {
            List<string> failedFiles = new List<string>();

            try
            {
                var inventorApp = GetInventorApplication();

                Documents docs = inventorApp.Documents;
                if (!Directory.Exists(directoryPath))
                {
                    Console.WriteLine($"Error: Directory not found -> {directoryPath}");
                    return false;
                }

                string originalPrefix = iProperties.GetValueOrDefault("originalPrefix", "");

                // Get all Inventor files, excluding unwanted folders
                var files = Directory.GetFiles(directoryPath, "*.*", SearchOption.AllDirectories)
                    .Where(f => f.IndexOf("OldVersions", StringComparison.OrdinalIgnoreCase) < 0 &&
                                f.IndexOf("BOUGHT OUT", StringComparison.OrdinalIgnoreCase) < 0 &&
                                f.IndexOf("ALLUSERSPROFILE", StringComparison.OrdinalIgnoreCase) < 0 &&
                                (f.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) ||
                                 f.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase)))
                    .Where(f => System.IO.Path.GetFileNameWithoutExtension(f)
                                 .StartsWith(originalPrefix, StringComparison.OrdinalIgnoreCase)) // <--- ADD THIS
                    .Select(f => new FileInfo(f))
                    .ToList();

                if (!files.Any())
                {
                    Console.WriteLine("No Inventor files found in directory.");
                    return false;
                }

                // Sort files - Parts first, then Assemblies, both in descending order
                var sortedFiles = files
                    .OrderByDescending(f => f.Extension == ".ipt") // Parts first
                    .ThenByDescending(f => f.Name) // Then by name descending
                    .ToList();

                Console.WriteLine($"\nFound {sortedFiles.Count} files to process (excluding OldVersions folders):");
                string partPrefix = iProperties.GetValueOrDefault("partPrefix", "");

                inventorApp.SilentOperation = true;
                inventorApp.Visible = false; // Hide Inventor window during processing

                foreach (var file in sortedFiles)
                {
                    string filePath = file.FullName;
                    bool fileUpdated = true;
                    Console.WriteLine($"\nProcessing file: {filePath}");

                    Document? inventorDoc = null;
                    try
                    {
                        // Open document with full access
                        inventorDoc = docs.Open(filePath, true);
                        PropertySets propSets = inventorDoc.PropertySets;

                        // Update properties in all property sets
                        foreach (var entry in iProperties)
                        {
                            if (entry.Key == "partPrefix" || entry.Key == "originalPrefix") continue;
                            bool propertyUpdated = false;

                            // Try to update in each property set
                            foreach (PropertySet propSet in propSets)
                            {
                                try
                                {
                                    if (propSet.Name == "Design Tracking Properties" ||
                                        propSet.Name == "Summary Information" ||
                                        propSet.Name == "Project Information" ||
                                        propSet.Name == "Inventor Document Summary Information")
                                    {
                                        Property? prop = null;
                                        try
                                        {
                                            prop = propSet[entry.Key];
                                        }
                                        catch
                                        {
                                            // Property doesn't exist in this set, try next set
                                            continue;
                                        }

                                        if (prop != null)
                                        {
                                            prop.Value = entry.Value;
                                            Console.WriteLine($"✅ Updated {entry.Key} = {entry.Value} in {propSet.Name}");
                                            propertyUpdated = true;
                                            break;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Warning: Could not update property {entry.Key} in {propSet.Name}: {ex.Message}");
                                }
                            }

                            if (!propertyUpdated)
                            {
                                Console.WriteLine($"❌ Failed to update property: {entry.Key}");
                                fileUpdated = false;
                            }
                        }

                        // Update Part Number if partPrefix is provided
                        if (!string.IsNullOrEmpty(partPrefix))
                        {
                            try
                            {
                                PropertySet designTrackingProps = propSets["Design Tracking Properties"];
                                Property partNumberProp = designTrackingProps["Part Number"];
                                string existingPartNumber = partNumberProp.Value?.ToString() ?? "";

                                string newPartNumber = existingPartNumber.Contains("_")
                                    ? $"{partPrefix}_{existingPartNumber[(existingPartNumber.IndexOf('_') + 1)..]}"
                                    : $"{partPrefix}_{existingPartNumber}";

                                partNumberProp.Value = newPartNumber;
                                Console.WriteLine($"✅ Updated: Part Number = {newPartNumber}");
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine($"❌ Failed to update Part Number: {e.Message}");
                                fileUpdated = false;
                            }
                        }

                        // Update and save
                        try
                        {
                            // Update document
                            inventorDoc.Update2();
                            Console.WriteLine($"🔄 Update triggered for: {filePath}");

                            // Update mass properties and rebuild document
                            if (inventorDoc is PartDocument partDoc)
                            {
                                try
                                {
                                    if (partDoc.ComponentDefinition.SurfaceBodies.Count == 0)
                                    {
                                        Console.WriteLine($"Skipping mass properties update for {filePath}: No solid bodies.");
                                    }
                                    else
                                    {
                                        // Update mass properties
                                        MassProperties massProps = partDoc.ComponentDefinition.MassProperties;
                                        massProps.Accuracy = 0; // Set to high accuracy
                                        // Force mass properties update by accessing properties
                                        double mass = massProps.Mass;
                                        double volume = massProps.Volume;
                                        double area = massProps.Area;

                                        partDoc.Rebuild();
                                        Console.WriteLine($"🔄 Mass properties updated and rebuild completed for part: {filePath}");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Warning: Could not update mass properties for part: {ex.Message}");
                                }
                            }
                            else if (inventorDoc is AssemblyDocument asmDoc)
                            {
                                try
                                {
                                    // Update mass properties
                                    MassProperties massProps = asmDoc.ComponentDefinition.MassProperties;
                                    massProps.Accuracy = 0; // Set to high accuracy
                                    // Force mass properties update by accessing properties
                                    double mass = massProps.Mass;
                                    double volume = massProps.Volume;
                                    double area = massProps.Area;

                                    asmDoc.Rebuild();
                                    Console.WriteLine($"🔄 Mass properties updated and rebuild completed for assembly: {filePath}");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Warning: Could not update mass properties for assembly: {ex.Message}");
                                }
                            }

                            inventorApp.ActiveView.Update();
                            inventorDoc.Save2(true);
                            Console.WriteLine($"💾 Save triggered for: {filePath}");
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"❌ Failed to update/save: {e.Message}");
                            fileUpdated = false;
                        }
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"❌ Error processing file: {filePath} -> {e.Message}");
                        fileUpdated = false;
                    }
                    finally
                    {
                        if (inventorDoc != null)
                        {
                            try
                            {
                                inventorDoc.Close(true);
                                Marshal.ReleaseComObject(inventorDoc);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine($"Error closing document: {e.Message}");
                            }
                        }
                    }

                    if (!fileUpdated)
                    {
                        failedFiles.Add(filePath);
                    }
                }

                // Log failed files
                if (failedFiles.Any())
                {
                    Console.WriteLine($"\n⚠️ {failedFiles.Count} files were NOT updated:");
                    foreach (string failedFile in failedFiles)
                    {
                        Console.WriteLine($" - {failedFile}");
                    }
                }

                return !failedFiles.Any();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error updating properties: {e.Message}");
                return false;
            }
            finally
            {
                // Cleanup Inventor and COM objects
                try
                {
                    if (_inventorApp != null)
                    {
                        // Close all remaining documents
                        while (_inventorApp.Documents.Count > 0)
                        {
                            try
                            {
                                Document doc = _inventorApp.Documents[1];
                                doc.Close(true);
                                Marshal.ReleaseComObject(doc);
                            }
                            catch (Exception ex)
                            {
                                Console.Error.WriteLine($"Error closing document: {ex.Message}");
                            }
                        }

                        // Quit Inventor and release COM object
                        _inventorApp.Quit();
                        Marshal.ReleaseComObject(_inventorApp);
                        _inventorApp = null;

                        // Force garbage collection to ensure COM objects are released
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                }
                catch (Exception ex)
                {
                    Console.Error.WriteLine($"Error during cleanup: {ex.Message}");
                }
            }
        }

        internal bool UpdateIpartsAndIassemblies(List<AssemblyUpdate> assemblyUpdates)
        {
            try
            {
                var inventorApp = GetInventorApplication();
                inventorApp.SilentOperation = true; // Suppress dialogs

                foreach (var update in assemblyUpdates)
                {
                    string assemblyFilePath = System.IO.Path.Combine("D:\\PROJECTS\\VECTOR\\3D Modelling\\TRANSFORMER\\WIP\\ABC099001\\MODEL", update.AssemblyFilePath);

                    AssemblyDocument? assemblyDoc = null;
                    try
                    {
                        assemblyDoc = (AssemblyDocument)inventorApp.Documents.Open(assemblyFilePath);
                        ComponentOccurrences occurrences = assemblyDoc.ComponentDefinition.Occurrences;

                        foreach (var (oldComponent, newComponent) in update.IpartsIassemblies)
                        {
                            ComponentOccurrence? occurrence = occurrences.Cast<ComponentOccurrence>()
                                .FirstOrDefault(o => o.Name.Equals(oldComponent, StringComparison.OrdinalIgnoreCase));

                            if (occurrence != null)
                            {
                                Console.WriteLine($"Processing component: {oldComponent}");

                                if (occurrence.Definition is PartComponentDefinition partDef && partDef.IsiPartMember)
                                {
                                    Console.WriteLine($"Changing iPart {oldComponent} to {newComponent}");
                                    try
                                    {
                                        // Parse the occurrence name to get the base name and instance number
                                        string[] parts = oldComponent.Split(':');
                                        if (parts.Length == 2)
                                        {
                                            string baseName = parts[0];
                                            if (int.TryParse(parts[1], out int instanceNumber))
                                            {
                                                // Find the specific occurrence in the assembly
                                                ComponentOccurrence? targetOccurrence = null;
                                                foreach (ComponentOccurrence occ in occurrences)
                                                {
                                                    if (occ.Name.StartsWith(baseName) && occ.Name.Contains(":" + instanceNumber))
                                                    {
                                                        targetOccurrence = occ;
                                                        break;
                                                    }
                                                }

                                                if (targetOccurrence != null)
                                                {
                                                    // Get the factory document path
                                                    string docPath = targetOccurrence.ReferencedDocumentDescriptor.FullDocumentName;
                                                    string? directoryPath = System.IO.Path.GetDirectoryName(docPath);
                                                    if (directoryPath == null)
                                                    {
                                                        throw new InvalidOperationException($"Could not get directory path for document: {docPath}");
                                                    }
                                                    string factoryPath = System.IO.Path.Combine(
                                                        directoryPath,
                                                        System.IO.Path.GetFileNameWithoutExtension(docPath).Split(':')[0] + ".ipt");

                                                    // Create the new member file path
                                                    string newMemberPath = System.IO.Path.Combine(
                                                        directoryPath,
                                                        newComponent + ".ipt");

                                                    // Replace the occurrence with the new member
                                                    targetOccurrence.Replace(newMemberPath, false);

                                                    Console.WriteLine($"Successfully replaced iPart instance {oldComponent} with {newComponent}");
                                                }
                                                else
                                                {
                                                    Console.WriteLine($"Could not find specific occurrence {oldComponent}");
                                                }
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Could not parse instance number from {oldComponent}");
                                            }
                                        }
                                        else
                                        {
                                            // If the component name doesn't have an instance number, try direct replacement
                                            string? refDocPath = occurrence.ReferencedDocumentDescriptor.FullDocumentName;
                                            string? refDirectoryPath = System.IO.Path.GetDirectoryName(refDocPath);
                                            if (refDirectoryPath == null)
                                            {
                                                throw new InvalidOperationException($"Could not get directory path for referenced document: {refDocPath}");
                                            }
                                            string newPath = System.IO.Path.Combine(
                                                refDirectoryPath,
                                                newComponent + ".ipt");

                                            if (System.IO.File.Exists(newPath))
                                            {
                                                occurrence.Replace(newPath, false);
                                                Console.WriteLine($"Successfully replaced iPart {oldComponent} with {newComponent}");
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Could not find iPart file: {newPath}");
                                            }
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine($"Error updating iPart: {e.Message}");
                                    }
                                }
                                else if (occurrence.Definition is AssemblyComponentDefinition asmDef && asmDef.IsiAssemblyMember)
                                {
                                    Console.WriteLine($"Changing iAssembly {oldComponent} to {newComponent}");
                                    try
                                    {
                                        // Parse the occurrence name to get the base name and instance number
                                        string[] parts = oldComponent.Split(':');
                                        if (parts.Length == 2)
                                        {
                                            string baseName = parts[0];
                                            if (int.TryParse(parts[1], out int instanceNumber))
                                            {
                                                // Find the specific occurrence in the assembly
                                                ComponentOccurrence? targetOccurrence = null;
                                                foreach (ComponentOccurrence occ in occurrences)
                                                {
                                                    if (occ.Name.StartsWith(baseName) && occ.Name.Contains(":" + instanceNumber))
                                                    {
                                                        targetOccurrence = occ;
                                                        break;
                                                    }
                                                }

                                                if (targetOccurrence != null)
                                                {
                                                    // Get the factory document path
                                                    string docPath = targetOccurrence.ReferencedDocumentDescriptor.FullDocumentName;
                                                    string? directoryPath = System.IO.Path.GetDirectoryName(docPath);
                                                    if (directoryPath == null)
                                                    {
                                                        throw new InvalidOperationException($"Could not get directory path for document: {docPath}");
                                                    }
                                                    string factoryPath = System.IO.Path.Combine(
                                                        directoryPath,
                                                        System.IO.Path.GetFileNameWithoutExtension(docPath).Split(':')[0] + ".iam");

                                                    // Create the new member file path
                                                    string newMemberPath = System.IO.Path.Combine(
                                                        directoryPath,
                                                        newComponent + ".iam");

                                                    // Replace the occurrence with the new member
                                                    targetOccurrence.Replace(newMemberPath, false);

                                                    Console.WriteLine($"Successfully replaced iAssembly instance {oldComponent} with {newComponent}");
                                                }
                                                else
                                                {
                                                    Console.WriteLine($"Could not find specific occurrence {oldComponent}");
                                                }
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Could not parse instance number from {oldComponent}");
                                            }
                                        }
                                        else
                                        {
                                            // If the component name doesn't have an instance number, try direct replacement
                                            string? refDocPath = occurrence.ReferencedDocumentDescriptor.FullDocumentName;
                                            string? refDirectoryPath = System.IO.Path.GetDirectoryName(refDocPath);
                                            if (refDirectoryPath == null)
                                            {
                                                throw new InvalidOperationException($"Could not get directory path for referenced document: {refDocPath}");
                                            }
                                            string newPath = System.IO.Path.Combine(
                                                refDirectoryPath,
                                                newComponent + ".iam");

                                            if (System.IO.File.Exists(newPath))
                                            {
                                                occurrence.Replace(newPath, false);
                                                Console.WriteLine($"Successfully replaced iAssembly {oldComponent} with {newComponent}");
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Could not find iAssembly file: {newPath}");
                                            }
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine($"Error updating iAssembly: {e.Message}");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($"Replacing normal component: {oldComponent} with {newComponent}");
                                    try
                                    {
                                        occurrence.Replace(newComponent, false);
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine($"Error replacing component: {e.Message}");
                                    }
                                }
                            }
                            else
                            {
                                Console.WriteLine($"Component {oldComponent} not found in {assemblyFilePath}");
                            }
                        }

                        try
                        {
                            assemblyDoc.Update();
                            inventorApp.ActiveView.Update();
                            assemblyDoc.Save2(true); // Save with Yes to All, suppress dialogs
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine($"Error saving assembly: {e.Message}");
                        }
                    }
                    finally
                    {
                        if (assemblyDoc != null)
                        {
                            try
                            {
                                assemblyDoc.Close(true); // Close and save
                                Marshal.ReleaseComObject(assemblyDoc);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine($"Error closing assembly document: {e.Message}");
                            }
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error updating iParts/iAssemblies: {ex.Message}");
                return false;
            }
            finally
            {
                if (_inventorApp != null)
                {
                    _inventorApp.SilentOperation = false; // Reset after operation
                    // Cleanup Inventor and COM objects
                    try
                    {
                        if (_inventorApp != null)
                        {
                            // Close all remaining documents
                            while (_inventorApp.Documents.Count > 0)
                            {
                                try
                                {
                                    Document doc = _inventorApp.Documents[1];
                                    doc.Close(true);
                                    Marshal.ReleaseComObject(doc);
                                }
                                catch (Exception ex)
                                {
                                    Console.Error.WriteLine($"Error closing document: {ex.Message}");
                                }
                            }

                            // Quit Inventor and release COM object
                            _inventorApp.Quit();
                            Marshal.ReleaseComObject(_inventorApp);
                            _inventorApp = null;

                            // Force garbage collection to ensure COM objects are released
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Error during cleanup: {ex.Message}");
                    }
                }
            }
        }

        public void ChangeIPartMember(string filePath, string memberName)
        {
            try
            {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                Document doc = _inventorApp.Documents.Open(filePath);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                if (!(doc is PartDocument partDoc))
                {
                    throw new Exception("The specified file is not a part document.");
                }

                PartComponentDefinition partDef = partDoc.ComponentDefinition;

                try
                {
                    // Check if this is an iPart member
                    if (partDef.IsiPartMember)
                    {
                        iPartMember member = partDef.iPartMember;

                        // Fixed method name from ChangeToMember to ChangeToRow
                        member.ChangeRow(memberName);
                        partDoc.Save();
                        Console.WriteLine($"Successfully changed to iPart member {memberName}");
                    }
                    else
                    {
                        throw new Exception("This part is not an iPart member.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error accessing iPart: {ex.Message}");
                    throw;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error changing iPart member: {ex.Message}");
            }
        }

        public void ListAllIPartMembers(string filePath)
        {
            try
            {
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                Document doc = _inventorApp.Documents.Open(filePath);
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                if (!(doc is PartDocument partDoc))
                {
                    throw new Exception("The specified file is not a part document.");
                }

                PartComponentDefinition partDef = partDoc.ComponentDefinition;

                try
                {
                    if (partDef.IsiPartFactory)
                    {
                        // This is an iPart factory
                        iPartFactory factory = partDef.iPartFactory;

                        Console.WriteLine("Available iPart members in factory:");
                        for (int i = 0; i < factory.TableRows.Count; i++)
                        {
                            Console.WriteLine($"- {factory.TableRows[i].MemberName}");
                        }
                    }
                    else if (partDef.IsiPartMember)
                    {
                        // This is an iPart member
#pragma warning disable CS8600 // Converting null literal or possible null value to non-nullable type.
                        string factoryPath = partDoc.PropertySets["Design Tracking Properties"]["Catalog Web Link"].Value.ToString();
#pragma warning restore CS8600 // Converting null literal or possible null value to non-nullable type.
                        Document factoryDoc = _inventorApp.Documents.Open(factoryPath);
                        PartDocument factoryPartDoc = (PartDocument)factoryDoc;

                        iPartFactory factory = factoryPartDoc.ComponentDefinition.iPartFactory;

                        Console.WriteLine("Available iPart members in factory:");
                        for (int i = 0; i < factory.TableRows.Count; i++)
                        {
                            Console.WriteLine($"- {factory.TableRows[i].MemberName}");
                        }

                        factoryDoc.Close();
                    }
                    else
                    {
                        Console.WriteLine("This part is not an iPart factory or member.");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error accessing iPart information: {ex.Message}");
                    throw;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error listing iPart members: {ex.Message}");
            }
        }

        private int GetIpartRowIndex(iPartFactory factory, string newComponentName)
        {
            for (int i = 0; i < factory.TableRows.Count; i++) // Fixed loop condition
            {
                iPartTableRow row = factory.TableRows[i];

                if (row.MemberName.Equals(newComponentName.Trim(), StringComparison.OrdinalIgnoreCase)) // Added Trim()
                {
                    return i; // Return correct row index
                }
            }
            return -1; // Not found
        }

        private int GetIAssemblyRowIndex(AssemblyComponentDefinition asmDef, string newComponentName)
        {
            iAssemblyFactory factory = asmDef.iAssemblyFactory;

            for (int i = 0; i < factory.TableRows.Count; i++)  // Fixed loop condition
            {
                iAssemblyTableRow row = factory.TableRows[i];

                if (row.MemberName.Equals(newComponentName, StringComparison.OrdinalIgnoreCase))
                {
                    return i; // Return correct row index
                }
            }
            return -1; // Not found
        }

        public bool UpdateModelStateAndRepresentations(List<ModelStateUpdate> updates)
        {
            try
            {
                var inventorApp = GetInventorApplication();
                inventorApp.SilentOperation = true; // Suppress dialogs

                foreach (var update in updates)
                {
                    string assemblyFilePath = System.IO.Path.Combine("D:\\PROJECTS\\VECTOR\\3D Modelling\\TRANSFORMER\\WIP\\ABC099001\\MODEL", update.AssemblyFilePath + ".iam");

                    if (!System.IO.File.Exists(assemblyFilePath))
                    {
                        Console.WriteLine($"Assembly file not found: {assemblyFilePath}");
                        continue;
                    }

                    Document? doc = null;
                    try
                    {
                        doc = inventorApp.Documents.Open(assemblyFilePath, true); // Open with full access

                        if (doc is AssemblyDocument asmDoc)
                        {
                            // Update Model State if specified
                            if (!string.IsNullOrEmpty(update.ModelState))
                            {
                                try
                                {
                                    // Get the model states
                                    ModelStates modelStates = asmDoc.ComponentDefinition.ModelStates;

                                    // Find and activate the specified model state
                                    foreach (ModelState state in modelStates)
                                    {
                                        if (state.Name.Equals(update.ModelState, StringComparison.OrdinalIgnoreCase))
                                        {
                                            Console.WriteLine($"Activating model state: {state.Name}");
                                            state.Activate();
                                            break;
                                        }
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error activating model state: {ex.Message}");
                                }
                            }

                            // Update Representation if specified
                            if (!string.IsNullOrEmpty(update.Representations))
                            {
                                try
                                {
                                    // Get the representations manager
                                    RepresentationsManager repManager = asmDoc.ComponentDefinition.RepresentationsManager;

                                    // First check in Design View Representations
                                    bool representationFound = false;

                                    foreach (DesignViewRepresentation rep in repManager.DesignViewRepresentations)
                                    {
                                        if (rep.Name.Equals(update.Representations, StringComparison.OrdinalIgnoreCase))
                                        {
                                            Console.WriteLine($"Activating design view representation: {rep.Name}");
                                            rep.Activate();
                                            representationFound = true;
                                            break;
                                        }
                                    }

                                    // If not found in Design Views, check in Positional Representations
                                    if (!representationFound)
                                    {
                                        foreach (PositionalRepresentation rep in repManager.PositionalRepresentations)
                                        {
                                            if (rep.Name.Equals(update.Representations, StringComparison.OrdinalIgnoreCase))
                                            {
                                                Console.WriteLine($"Activating positional representation: {rep.Name}");
                                                rep.Activate();
                                                representationFound = true;
                                                break;
                                            }
                                        }
                                    }

                                    // If not found in Positional, check in Level of Detail Representations
                                    if (!representationFound)
                                    {
                                        foreach (LevelOfDetailRepresentation rep in repManager.LevelOfDetailRepresentations)
                                        {
                                            if (rep.Name.Equals(update.Representations, StringComparison.OrdinalIgnoreCase))
                                            {
                                                Console.WriteLine($"Activating level of detail representation: {rep.Name}");
                                                rep.Activate();
                                                representationFound = true;
                                                break;
                                            }
                                        }
                                    }

                                    if (!representationFound)
                                    {
                                        Console.WriteLine($"Warning: Could not find representation named '{update.Representations}'");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error activating representation: {ex.Message}");
                                }
                            }

                            // Make sure to update the view and save
                            inventorApp.ActiveView.Update();
                            asmDoc.Save2(true); // Save with Yes to All, suppress dialogs
                        }
                        else
                        {
                            Console.WriteLine($"Document is not an assembly: {assemblyFilePath}");
                        }
                    }
                    finally
                    {
                        if (doc != null)
                        {
                            try
                            {
                                doc.Close(true); // Close and save
                                Marshal.ReleaseComObject(doc);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine($"Error closing document: {e.Message}");
                            }
                        }
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error updating model states and representations: {ex.Message}");
                return false;
            }
            finally
            {
                if (_inventorApp != null)
                {
                    _inventorApp.SilentOperation = false; // Reset after operation
                    // Cleanup Inventor and COM objects
                    try
                    {
                        if (_inventorApp != null)
                        {
                            // Close all remaining documents
                            while (_inventorApp.Documents.Count > 0)
                            {
                                try
                                {
                                    Document doc = _inventorApp.Documents[1];
                                    doc.Close(true);
                                    Marshal.ReleaseComObject(doc);
                                }
                                catch (Exception ex)
                                {
                                    Console.Error.WriteLine($"Error closing document: {ex.Message}");
                                }
                            }

                            // Quit Inventor and release COM object
                            _inventorApp.Quit();
                            Marshal.ReleaseComObject(_inventorApp);
                            _inventorApp = null;

                            // Force garbage collection to ensure COM objects are released
                            GC.Collect();
                            GC.WaitForPendingFinalizers();
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine($"Error during cleanup: {ex.Message}");
                    }
                }
            }
        }
        public bool DesignAssistRename(string drawingsPath, List<string> assemblyList, string partPrefix)
        {
            var warnings = new List<string>();

            try
            {
                var inventorApp = GetInventorApplication();
                inventorApp.SilentOperation = true; // Suppress dialogs
                inventorApp.Visible = false; // Hide Inventor window

                foreach (var mainAssembly in assemblyList)
                {
                    string mainAssemblyPath = System.IO.Path.Combine(drawingsPath, mainAssembly);
                    if (!System.IO.File.Exists(mainAssemblyPath))
                    {
                        warnings.Add($"Main assembly file not found: {mainAssemblyPath}");
                        continue;
                    }

                    Console.WriteLine($"Processing main assembly: {mainAssemblyPath}");
                    AssemblyDocument? asmDoc = null;

                    try
                    {
                        asmDoc = (AssemblyDocument)inventorApp.Documents.Open(mainAssemblyPath, true); // Open with full access
                        var occurrences = asmDoc.ComponentDefinition.Occurrences;
                        var renameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                        foreach (ComponentOccurrence occ in occurrences)
                        {
                            try
                            {
                                string refPath = occ.ReferencedDocumentDescriptor.FullDocumentName;
                                string partNumber = GetPartNumberFromFile(refPath);
                                if (string.IsNullOrEmpty(partNumber)) continue;

                                // Only rename files whose part number already starts with the provided prefix
                                if (!partNumber.StartsWith(partPrefix, StringComparison.OrdinalIgnoreCase))
                                    continue;

                                string ext = System.IO.Path.GetExtension(refPath);
                                string dir = System.IO.Path.GetDirectoryName(refPath)!;

                                // Remove old prefix if present and add the new one
                                string suffix = partNumber.Substring(partPrefix.Length).TrimStart('_');
                                string newName = $"{partPrefix}_{suffix}";
                                string newPath = System.IO.Path.Combine(dir, newName + ext);

                                if (!System.IO.File.Exists(newPath) && !renameMap.ContainsKey(refPath))
                                {
                                    renameMap.Add(refPath, newPath);
                                }
                            }
                            catch (Exception ex)
                            {
                                warnings.Add($"Warning: Could not process occurrence: {ex.Message}");
                            }
                        }

                        foreach (var kvp in renameMap)
                        {
                            string oldPath = kvp.Key;
                            string newPath = kvp.Value;

                            try
                            {
                                System.IO.File.Move(oldPath, newPath);
                                Console.WriteLine($"Renamed file: {oldPath} -> {newPath}");

                                foreach (ComponentOccurrence occ in occurrences)
                                {
                                    if (occ.ReferencedDocumentDescriptor.FullDocumentName.Equals(oldPath, StringComparison.OrdinalIgnoreCase))
                                    {
                                        occ.Replace(newPath, false);
                                        Console.WriteLine($"Updated reference: {oldPath} -> {newPath}");
                                    }
                                }
                            }
                            catch (Exception ex)
                            {
                                warnings.Add($"Failed to rename/update {oldPath}: {ex.Message}");
                            }
                        }

                        asmDoc.Update();
                        inventorApp.ActiveView.Update();
                        asmDoc.Save2(true); // Save with Yes to All, suppress dialogs
                        Console.WriteLine($"Saved assembly: {mainAssemblyPath}");
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Error processing assembly {mainAssemblyPath}: {ex.Message}");
                    }
                    finally
                    {
                        if (asmDoc != null)
                        {
                            try
                            {
                                asmDoc.Close(true);
                                Marshal.ReleaseComObject(asmDoc);
                                asmDoc = null;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error closing assembly document: {ex.Message}");
                            }
                        }
                    }

                    // Optionally rename the main assembly file itself
                    string mainPartNumber = GetPartNumberFromFile(mainAssemblyPath);
                    if (!string.IsNullOrEmpty(mainPartNumber) &&
                        mainPartNumber.StartsWith(partPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        string suffix = mainPartNumber.Substring(partPrefix.Length).TrimStart('_');
                        string newMainName = $"{partPrefix}_{suffix}";
                        string mainExt = System.IO.Path.GetExtension(mainAssemblyPath);
                        string mainNewPath = System.IO.Path.Combine(drawingsPath, newMainName + mainExt);

                        if (!System.IO.File.Exists(mainNewPath))
                        {
                            try
                            {
                                System.IO.File.Move(mainAssemblyPath, mainNewPath);
                                Console.WriteLine($"Renamed main assembly: {mainAssemblyPath} -> {mainNewPath}");
                            }
                            catch (Exception ex)
                            {
                                warnings.Add($"Failed to rename main assembly: {ex.Message}");
                            }
                        }
                        else
                        {
                            warnings.Add($"Target main assembly file already exists: {mainNewPath}");
                        }
                    }
                }

                if (warnings.Count > 0)
                {
                    foreach (var w in warnings)
                        Console.WriteLine(w);
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Inventor API error: {ex.Message}");
                return false;
            }
            finally
            {
                if (_inventorApp != null)
                {
                    _inventorApp.SilentOperation = false; // Reset after operation

                    // Close all remaining documents
                    while (_inventorApp.Documents.Count > 0)
                    {
                        try
                        {
                            Document doc = _inventorApp.Documents[1];
                            doc.Close(true);
                            Marshal.ReleaseComObject(doc);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error closing document: {ex.Message}");
                        }
                    }

                    try
                    {
                        _inventorApp.Quit();
                        Marshal.ReleaseComObject(_inventorApp);
                        _inventorApp = null;
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error during Inventor cleanup: {ex.Message}");
                    }
                }
            }
        }

        // Helper method to get the Part Number property from a file
        private string GetPartNumberFromFile(string filePath)
        {
            try
            {
                Type? inventorType = Type.GetTypeFromProgID("Inventor.Application");
                if (inventorType == null)
                    throw new Exception("Inventor is not installed.");
                dynamic? inventorApp = Activator.CreateInstance(inventorType);
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                inventorApp.Visible = false;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                dynamic doc = inventorApp.Documents.Open(filePath, true);
                string partNumber = "";
                try
                {
                    var propSets = doc.PropertySets;
                    var designProps = propSets["Design Tracking Properties"];
                    partNumber = designProps["Part Number"].Value.ToString();
                }
                finally
                {
                    doc.Close();
                    inventorApp.Quit();
                }
                return partNumber ?? "";
            }
            catch
            {
                return "";
            }
        }
    }
}

