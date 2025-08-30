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

        public bool DesignAssistRename(string drawingsPath, string newPrefix, List<string>? assemblyList = null)
        {
            var warnings = new List<string>();
            var processedAssemblies = new List<AssemblyDocument>();

            try
            {
                var inventorApp = GetInventorApplication();
                inventorApp.SilentOperation = true;
                inventorApp.Visible = false;

                // Auto-discover assembly files if no list provided
                List<string> assemblies;
                if (assemblyList == null || assemblyList.Count == 0)
                {
                    Console.WriteLine($"Auto-discovering assembly files in: {drawingsPath}");
                    assemblies = DiscoverAssemblyFiles(drawingsPath);

                    if (assemblies.Count == 0)
                    {
                        warnings.Add($"No assembly files found in path: {drawingsPath}");
                        return false;
                    }

                    Console.WriteLine($"Found {assemblies.Count} assembly files to process:");
                    assemblies.ForEach(a => Console.WriteLine($"  - {a}"));
                }
                else
                {
                    assemblies = assemblyList;
                }

                // Global rename tracking with conflict resolution
                var globalRenameMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // Phase 1: Open all assemblies and build global rename map
                foreach (var mainAssembly in assemblies)
                {
                    string mainAssemblyPath = System.IO.Path.IsPathRooted(mainAssembly)
                        ? mainAssembly
                        : System.IO.Path.Combine(drawingsPath, mainAssembly);

                    if (!System.IO.File.Exists(mainAssemblyPath))
                    {
                        warnings.Add($"Main assembly file not found: {mainAssemblyPath}");
                        continue;
                    }

                    Console.WriteLine($"Opening assembly: {mainAssemblyPath}");
                    AssemblyDocument? asmDoc = null;

                    try
                    {
                        asmDoc = (AssemblyDocument)inventorApp.Documents.Open(mainAssemblyPath, true);
                        processedAssemblies.Add(asmDoc);

                        var occurrences = asmDoc.ComponentDefinition.Occurrences;

                        // Build rename map for this assembly
                        foreach (ComponentOccurrence occ in occurrences)
                        {
                            try
                            {
                                string refPath = occ.ReferencedDocumentDescriptor.FullDocumentName;
                                string fileName = System.IO.Path.GetFileNameWithoutExtension(refPath);

                                // *** KEY FIX 1: Skip Content Center files ***
                                if (IsContentCenterFile(refPath))
                                {
                                    Console.WriteLine($"Skipping Content Center file: {fileName}");
                                    continue;
                                }

                                // *** KEY FIX 2: Only rename files that match the part prefix pattern ***
                                if (!ShouldRename(occ, newPrefix))
                                {
                                    Console.WriteLine($"Skipping file (doesn't match prefix pattern): {fileName}");
                                    continue;
                                }

                                // Skip if already starts with new prefix
                                if (fileName.StartsWith(newPrefix, StringComparison.OrdinalIgnoreCase))
                                    continue;

                                string ext = System.IO.Path.GetExtension(refPath);
                                string dir = System.IO.Path.GetDirectoryName(refPath)!;

                                // Generate new name with conflict resolution
                                string newFileName = GenerateUniqueFileName(fileName, newPrefix, ext, dir, usedNames);
                                string newPath = System.IO.Path.Combine(dir, newFileName);

                                if (!globalRenameMap.ContainsKey(refPath))
                                {
                                    globalRenameMap.Add(refPath, newPath);
                                    usedNames.Add(newFileName);
                                }
                            }
                            catch (Exception ex)
                            {
                                warnings.Add($"Warning: Could not process occurrence in {mainAssemblyPath}: {ex.Message}");
                            }
                        }

                        // Handle main assembly renaming
                        string mainFileName = System.IO.Path.GetFileNameWithoutExtension(mainAssemblyPath);
                        if (ShouldRenameAssembly(asmDoc, newPrefix) && !mainFileName.StartsWith(newPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            string mainExt = System.IO.Path.GetExtension(mainAssemblyPath);
                            string newMainFileName = GenerateUniqueFileName(mainFileName, newPrefix, mainExt, drawingsPath, usedNames);
                            string mainNewPath = System.IO.Path.Combine(drawingsPath, newMainFileName);

                            if (!globalRenameMap.ContainsKey(mainAssemblyPath))
                            {
                                globalRenameMap.Add(mainAssemblyPath, mainNewPath);
                                usedNames.Add(newMainFileName);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Error opening assembly {mainAssemblyPath}: {ex.Message}");
                    }
                }

                Console.WriteLine($"Total files to rename: {globalRenameMap.Count}");

                // Phase 2: Close all assemblies before renaming files
                foreach (var asmDoc in processedAssemblies)
                {
                    try
                    {
                        Console.WriteLine($"Closing assembly: {asmDoc.FullFileName}");
                        asmDoc.Close(false); // Don't save yet
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Error closing assembly {asmDoc.FullFileName}: {ex.Message}");
                    }
                }
                processedAssemblies.Clear();

                // Force garbage collection to release COM objects
                GC.Collect();
                GC.WaitForPendingFinalizers();

                // Wait a moment for file locks to be released
                Thread.Sleep(5000);

                // Phase 3: Perform file renaming
                var successfulRenames = new Dictionary<string, string>();
                foreach (var kvp in globalRenameMap)
                {
                    string oldPath = kvp.Key;
                    string newPath = kvp.Value;

                    try
                    {
                        if (System.IO.File.Exists(newPath))
                        {
                            warnings.Add($"Target file already exists, skipping: {newPath}");
                            continue;
                        }

                        System.IO.File.Move(oldPath, newPath);
                        successfulRenames.Add(oldPath, newPath);
                        Console.WriteLine($"Renamed file: {System.IO.Path.GetFileName(oldPath)} -> {System.IO.Path.GetFileName(newPath)}");
                    }
                    catch (UnauthorizedAccessException)
                    {
                        warnings.Add($"File is locked, cannot rename: {oldPath}");
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Failed to rename {oldPath}: {ex.Message}");
                    }
                }

                // Phase 3.5: Update derived part links before assemblies
                Console.WriteLine("=== Phase 3.5: Updating Derived Part Links ===");
                var allIptFiles = Directory.GetFiles(drawingsPath, "*.ipt", SearchOption.AllDirectories);
                foreach (var iptPath in allIptFiles)
                {
                    string currentPath = successfulRenames.ContainsKey(iptPath) ? successfulRenames[iptPath] : iptPath;
                    if (!System.IO.File.Exists(currentPath))
                        continue;
                    PartDocument? partDoc = null;
                    try
                    {
                        partDoc = (PartDocument)inventorApp.Documents.Open(currentPath, false);
                        bool updated = false;
                        var derivedList = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Cast<DerivedPartComponent>().ToList();
                        foreach (DerivedPartComponent derived in derivedList)
                        {
                            string basePath = derived.ReferencedDocumentDescriptor.FullDocumentName;
                            if (successfulRenames.ContainsKey(basePath))
                            {
                                string newBasePath = successfulRenames[basePath];

                                // Check if new base file exists and is not locked
                                if (!System.IO.File.Exists(newBasePath))
                                {
                                    Console.WriteLine($"New base file does not exist: {newBasePath}");
                                    continue;
                                }
                                if (IsFileLocked(newBasePath))
                                {
                                    Console.WriteLine($"New base file is locked: {newBasePath}");
                                    continue;
                                }

                                int retryCount = 0;
                                bool recreated = false;
                                while (retryCount < 3 && !recreated)
                                {
                                    try
                                    {
                                        if (derived.Definition is DerivedPartUniformScaleDef def)
                                        {
                                            var scale = def.ScaleFactor;
                                            derived.Delete();
                                            var newDefObj = partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.CreateDefinition(newBasePath);
                                            if (newDefObj is DerivedPartUniformScaleDef newDef)
                                            {
                                                newDef.ScaleFactor = scale;
                                                partDoc.ComponentDefinition.ReferenceComponents.DerivedPartComponents.Add((DerivedPartDefinition)newDef);
                                                updated = true;
                                                Console.WriteLine($"Recreated derived part link in {System.IO.Path.GetFileName(currentPath)}: {System.IO.Path.GetFileName(basePath)} -> {System.IO.Path.GetFileName(newBasePath)}");
                                            }
                                            else
                                            {
                                                Console.WriteLine($"Failed to create new derived part definition for {System.IO.Path.GetFileName(currentPath)}");
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine($"Unsupported derived part definition type ({derived.Definition.GetType().Name}) in {System.IO.Path.GetFileName(currentPath)}. Skipping.");
                                            continue;
                                        }
                                        Console.WriteLine($"Derived feature type: {derived.Definition.GetType().FullName} in {System.IO.Path.GetFileName(currentPath)}");
                                        recreated = true;
                                    }
                                    catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                                    {
                                        retryCount++;
                                        if (retryCount < 3)
                                        {
                                            Console.WriteLine($"E_FAIL recreating derived part link in {System.IO.Path.GetFileName(currentPath)} (attempt {retryCount}/3), saving and reopening part, retrying...");
                                            // Save and close, then reopen the part document
                                            try
                                            {
                                                partDoc.Save();
                                                partDoc.Close(false);
                                                Marshal.ReleaseComObject(partDoc);
                                                Thread.Sleep(500);
                                                partDoc = (PartDocument)inventorApp.Documents.Open(currentPath, false);
                                            }
                                            catch (Exception reopenEx)
                                            {
                                                Console.WriteLine($"Error saving/reopening part document: {reopenEx.Message}");
                                                break;
                                            }
                                            continue;
                                        }
                                        Console.WriteLine($"Failed to recreate derived part link in {System.IO.Path.GetFileName(currentPath)} after 3 attempts: {comEx.Message}");
                                        break;
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine($"Failed to recreate derived part link in {System.IO.Path.GetFileName(currentPath)}: {ex.Message}");
                                        break;
                                    }
                                }
                            }
                            else if (derived.ReferencedDocumentDescriptor.ReferenceMissing)
                            {
                                Console.WriteLine($"Skipped unresolved derived feature in {System.IO.Path.GetFileName(currentPath)} (no new base path found)");
                                continue;
                            }
                        }
                        if (updated)
                        {
                            partDoc.Update();
                            partDoc.Save();
                            Console.WriteLine($"Saved derived part: {currentPath}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing derived part {currentPath}: {ex.Message}");
                    }
                    finally
                    {
                        if (partDoc != null)
                        {
                            try
                            {
                                partDoc.Close(false);
                                Marshal.ReleaseComObject(partDoc);
                            }
                            catch { }
                        }
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        Thread.Sleep(200);
                    }
                }

                // Phase 4: Update references in assemblies BEFORE releasing COM objects
                Console.WriteLine("=== Phase 4: Updating References ===");

                foreach (var mainAssembly in assemblies)
                {
                    string originalPath = System.IO.Path.IsPathRooted(mainAssembly)
                        ? mainAssembly
                        : System.IO.Path.Combine(drawingsPath, mainAssembly);

                    // Determine the current path (may have been renamed)
                    string currentPath = successfulRenames.ContainsKey(originalPath)
                        ? successfulRenames[originalPath]
                        : originalPath;

                    if (!System.IO.File.Exists(currentPath))
                    {
                        warnings.Add($"Assembly not found after rename: {currentPath}");
                        continue;
                    }

                    // Wait and ensure file is not locked
                    if (IsFileLocked(currentPath))
                    {
                        Console.WriteLine($"Waiting for file lock to clear: {currentPath}");
                        Thread.Sleep(3000);
                        if (IsFileLocked(currentPath))
                        {
                            warnings.Add($"File still locked, skipping: {currentPath}");
                            continue;
                        }
                    }

                    AssemblyDocument? asmDoc = null;
                    try
                    {
                        Console.WriteLine($"Opening for reference update: {currentPath}");
                        asmDoc = (AssemblyDocument)inventorApp.Documents.Open(currentPath, false);
                        Thread.Sleep(1000); // Give Inventor time to load the document

                        bool referencesUpdated = false;
                        var failedUpdates = new List<string>();

                        var occurrences = asmDoc.ComponentDefinition.Occurrences;
                        Thread.Sleep(500); // Give Inventor time to process occurrences

                        // Create a snapshot of occurrences to avoid collection modification issues
                        var occurrenceData = new List<(ComponentOccurrence occ, string oldPath, string newPath)>();
                        foreach (ComponentOccurrence occ in occurrences)
                        {
                            try
                            {
                                string currentRefPath = occ.ReferencedDocumentDescriptor.FullDocumentName;
                                if (successfulRenames.ContainsKey(currentRefPath))
                                {
                                    string newRefPath = successfulRenames[currentRefPath];
                                    occurrenceData.Add((occ, currentRefPath, newRefPath));
                                }
                            }
                            catch (Exception ex)
                            {
                                warnings.Add($"Error reading occurrence reference: {ex.Message}");
                            }
                        }

                        // Process the updates
                        foreach (var (occ, oldPath, newPath) in occurrenceData)
                        {
                            int retryCount = 0;
                            bool updated = false;
                            while (retryCount < 3 && !updated)
                            {
                                try
                                {
                                    if (!System.IO.File.Exists(newPath))
                                    {
                                        failedUpdates.Add($"Target file doesn't exist: {System.IO.Path.GetFileName(newPath)}");
                                        break;
                                    }
                                    // Skip suppressed or unresolved occurrences
                                    if (occ.Suppressed || occ.ReferencedDocumentDescriptor.ReferenceMissing)
                                    {
                                        failedUpdates.Add($"Skipped suppressed or unresolved occurrence: {occ.Name} in {asmDoc.DisplayName}");
                                        break;
                                    }
                                    // Only replace if the reference is not already correct
                                    if (!string.Equals(occ.ReferencedDocumentDescriptor.FullDocumentName, newPath, StringComparison.OrdinalIgnoreCase))
                                    {
                                        occ.Replace(newPath, false);
                                        referencesUpdated = true;
                                        Console.WriteLine($"Updated reference: {System.IO.Path.GetFileName(oldPath)} -> {System.IO.Path.GetFileName(newPath)} in {asmDoc.DisplayName} (Occurrence: {occ.Name})");
                                    }
                                    updated = true;
                                }
                                catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                                {
                                    retryCount++;
                                    if (retryCount < 3)
                                    {
                                        Thread.Sleep(500);
                                        continue;
                                    }
                                    failedUpdates.Add($"Failed to update: {System.IO.Path.GetFileName(oldPath)} in {asmDoc.DisplayName} (Occurrence: {occ.Name}): {comEx.Message} (E_FAIL after 3 attempts)");
                                    break;
                                }
                                catch (Exception ex)
                                {
                                    failedUpdates.Add($"Failed to update: {System.IO.Path.GetFileName(oldPath)} in {asmDoc.DisplayName} (Occurrence: {occ.Name}): {ex.Message}");
                                    break;
                                }
                            }
                        }

                        // Report failed updates
                        if (failedUpdates.Count > 0)
                        {
                            warnings.AddRange(failedUpdates);
                        }

                        // Update and save if we made changes
                        if (referencesUpdated)
                        {
                            try
                            {
                                asmDoc.Update();
                                Thread.Sleep(1000); // Give Inventor time to update
                                asmDoc.Save();
                                Thread.Sleep(1000); // Give Inventor time to save
                                Console.WriteLine($"Saved assembly: {currentPath}");
                            }
                            catch (Exception updateEx)
                            {
                                warnings.Add($"Failed to update/save assembly {currentPath}: {updateEx.Message}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Error processing assembly {currentPath}: {ex.Message}");
                    }
                    finally
                    {
                        if (asmDoc != null)
                        {
                            try
                            {
                                asmDoc.Close(false);
                                Marshal.ReleaseComObject(asmDoc);
                            }
                            catch { }
                        }
                        // Force garbage collection
                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                        GC.Collect();
                        Thread.Sleep(500); // Small delay between assemblies
                    }
                }

                if (warnings.Count > 0)
                {
                    Console.WriteLine("\n=== WARNINGS ===");
                    foreach (var w in warnings)
                        Console.WriteLine(w);
                }

                Console.WriteLine($"\n=== OPERATION COMPLETE ===");
                Console.WriteLine($"Total files renamed: {successfulRenames.Count}");
                Console.WriteLine($"Total assemblies processed: {assemblies.Count}");

                return successfulRenames.Count > 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Inventor API error: {ex.Message}");
                return false;
            }
            finally
            {
                // Cleanup
                foreach (var doc in processedAssemblies)
                {
                    try
                    {
                        doc.Close(false);
                        Marshal.ReleaseComObject(doc);
                    }
                    catch { }
                }

                CleanupInventorApp();
                GC.Collect();
            }
        }

        /// <summary>
        /// Analyzes what files would be renamed without performing the actual rename operation
        /// </summary>
        public object AnalyzeDesignAssistRename(string drawingsPath, string newPrefix, List<string>? assemblyList = null)
        {
            var analysis = new
            {
                assembliesFound = 0,
                filesToRename = 0,
                filesSkipped = 0,
                contentCenterFiles = 0,
                alreadyCorrectPrefix = 0,
                noPartNumber = 0,
                assemblyList = new List<string>(),
                filesToRenameList = new List<object>(),
                filesSkippedList = new List<object>(),
                warnings = new List<string>()
            };

            var assembliesFound = 0;
            var filesToRename = 0;
            var contentCenterFiles = 0;
            var alreadyCorrectPrefix = 0;
            var noPartNumber = 0;
            var assemblyListResult = new List<string>();
            var filesToRenameList = new List<object>();
            var filesSkippedList = new List<object>();
            var warnings = new List<string>();

            try
            {
                var inventorApp = GetInventorApplication();
                inventorApp.SilentOperation = true;
                inventorApp.Visible = false;

                // Auto-discover assembly files if no list provided
                List<string> assemblies;
                if (assemblyList == null || assemblyList.Count == 0)
                {
                    Console.WriteLine($"Analyzing: Auto-discovering assembly files in: {drawingsPath}");
                    assemblies = DiscoverAssemblyFiles(drawingsPath);

                    if (assemblies.Count == 0)
                    {
                        warnings.Add($"No assembly files found in path: {drawingsPath}");
                        return new
                        {
                            assembliesFound = 0,
                            filesToRename = 0,
                            filesSkipped = 0,
                            contentCenterFiles = 0,
                            alreadyCorrectPrefix = 0,
                            noPartNumber = 0,
                            assemblyList = new List<string>(),
                            filesToRenameList = new List<object>(),
                            filesSkippedList = new List<object>(),
                            warnings = warnings
                        };
                    }

                    Console.WriteLine($"Found {assemblies.Count} assembly files to analyze:");
                    assemblies.ForEach(a => Console.WriteLine($"  - {a}"));
                }
                else
                {
                    assemblies = assemblyList;
                }

                assembliesFound = assemblies.Count;
                assemblyListResult = assemblies;

                // Analyze each assembly
                foreach (var mainAssembly in assemblies)
                {
                    string mainAssemblyPath = System.IO.Path.IsPathRooted(mainAssembly)
                        ? mainAssembly
                        : System.IO.Path.Combine(drawingsPath, mainAssembly);

                    if (!System.IO.File.Exists(mainAssemblyPath))
                    {
                        warnings.Add($"Main assembly file not found: {mainAssemblyPath}");
                        continue;
                    }

                    Console.WriteLine($"Analyzing assembly: {mainAssemblyPath}");
                    AssemblyDocument? asmDoc = null;

                    try
                    {
                        asmDoc = (AssemblyDocument)inventorApp.Documents.Open(mainAssemblyPath, true);

                        var occurrences = asmDoc.ComponentDefinition.Occurrences;

                        // Analyze occurrences in this assembly
                        foreach (ComponentOccurrence occ in occurrences)
                        {
                            try
                            {
                                string refPath = occ.ReferencedDocumentDescriptor.FullDocumentName;
                                string fileName = System.IO.Path.GetFileNameWithoutExtension(refPath);
                                string fileExt = System.IO.Path.GetExtension(refPath);

                                // Check if it's a Content Center file
                                if (IsContentCenterFile(refPath))
                                {
                                    contentCenterFiles++;
                                    filesSkippedList.Add(new
                                    {
                                        fileName = fileName + fileExt,
                                        reason = "Content Center file",
                                        fullPath = refPath
                                    });
                                    continue;
                                }

                                // Check if it already has the correct prefix
                                if (fileName.StartsWith(newPrefix, StringComparison.OrdinalIgnoreCase))
                                {
                                    alreadyCorrectPrefix++;
                                    filesSkippedList.Add(new
                                    {
                                        fileName = fileName + fileExt,
                                        reason = "Already has correct prefix",
                                        fullPath = refPath
                                    });
                                    continue;
                                }

                                // Check if it should be renamed based on part number
                                if (!ShouldRename(occ, newPrefix))
                                {
                                    noPartNumber++;
                                    filesSkippedList.Add(new
                                    {
                                        fileName = fileName + fileExt,
                                        reason = "Part number doesn't match prefix pattern",
                                        fullPath = refPath
                                    });
                                    continue;
                                }

                                // This file would be renamed
                                filesToRename++;
                                string newFileName = GenerateUniqueFileName(fileName, newPrefix, fileExt, System.IO.Path.GetDirectoryName(refPath)!, new HashSet<string>());
                                filesToRenameList.Add(new
                                {
                                    currentFileName = fileName + fileExt,
                                    newFileName = newFileName,
                                    fullPath = refPath,
                                    newFullPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(refPath)!, newFileName)
                                });
                            }
                            catch (Exception ex)
                            {
                                warnings.Add($"Warning: Could not analyze occurrence in {mainAssemblyPath}: {ex.Message}");
                            }
                        }

                        // Analyze main assembly
                        string mainFileName = System.IO.Path.GetFileNameWithoutExtension(mainAssemblyPath);
                        string mainExt = System.IO.Path.GetExtension(mainAssemblyPath);

                        if (ShouldRenameAssembly(asmDoc, newPrefix) && !mainFileName.StartsWith(newPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            filesToRename++;
                            string newMainFileName = GenerateUniqueFileName(mainFileName, newPrefix, mainExt, drawingsPath, new HashSet<string>());
                            filesToRenameList.Add(new
                            {
                                currentFileName = mainFileName + mainExt,
                                newFileName = newMainFileName,
                                fullPath = mainAssemblyPath,
                                newFullPath = System.IO.Path.Combine(drawingsPath, newMainFileName),
                                isMainAssembly = true
                            });
                        }
                    }
                    catch (Exception ex)
                    {
                        warnings.Add($"Error analyzing assembly {mainAssemblyPath}: {ex.Message}");
                    }
                    finally
                    {
                        if (asmDoc != null)
                        {
                            try
                            {
                                asmDoc.Close(false);
                                Marshal.ReleaseComObject(asmDoc);
                            }
                            catch { }
                        }
                    }
                }

                // Cleanup
                CleanupInventorApp();
                GC.Collect();

                return new
                {
                    assembliesFound,
                    filesToRename,
                    filesSkipped = contentCenterFiles + alreadyCorrectPrefix + noPartNumber,
                    contentCenterFiles,
                    alreadyCorrectPrefix,
                    noPartNumber,
                    assemblyList = assemblyListResult,
                    filesToRenameList,
                    filesSkippedList,
                    warnings
                };
            }
            catch (Exception ex)
            {
                warnings.Add($"Analysis error: {ex.Message}");
                return new
                {
                    assembliesFound = 0,
                    filesToRename = 0,
                    filesSkipped = 0,
                    contentCenterFiles = 0,
                    alreadyCorrectPrefix = 0,
                    noPartNumber = 0,
                    assemblyList = new List<string>(),
                    filesToRenameList = new List<object>(),
                    filesSkippedList = new List<object>(),
                    warnings = warnings
                };
            }
        }

        // *** HELPER METHODS TO ADD ***

        /// <summary>
        /// Enhanced method to try different approaches for replacing references with better error handling
        /// </summary>
        private bool TryReplaceReference(ComponentOccurrence occ, string oldPath, string newPath)
        {
            const int maxRetries = 3;
            const int retryDelayMs = 2000;

            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    // Verify the new file exists and is accessible
                    if (!System.IO.File.Exists(newPath))
                    {
                        Console.WriteLine($"Target file doesn't exist: {newPath}");
                        return false;
                    }

                    // Enhanced file lock checking with longer timeout
                    if (IsFileLocked(newPath, 5))
                    {
                        Console.WriteLine($"Target file is locked (attempt {attempt}/{maxRetries}): {newPath}");
                        if (attempt < maxRetries)
                        {
                            Thread.Sleep(retryDelayMs);
                            continue;
                        }
                        return false;
                    }

                    // Get the assembly document and start a transaction if possible
                    var asmDoc = (AssemblyDocument)occ.ContextDefinition.Document;

                    // Force update the document references first
                    try
                    {
                        ((Inventor.Document)asmDoc).Update();
                        Thread.Sleep(1000); // Brief pause after update
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Document update failed: {ex.Message}");
                    }

                    // Store the occurrence's properties before any operations
                    var transform = occ.Transformation;
                    var visible = occ.Visible;
                    var isSuppressed = occ.Suppressed;
                    var name = occ.Name;

                    // Method 1: Try direct replacement first
                    try
                    {
                        // Get the referenced document descriptor
                        var refDocDesc = occ.ReferencedDocumentDescriptor;
                        if (refDocDesc != null)
                        {
                            // Try to close the referenced document first
                            try
                            {
                                var refDoc = refDocDesc.ReferencedDocument;
                                if (refDoc != null)
                                {
                                    ((Inventor.Document)refDoc).Close(false);
                                    Thread.Sleep(1000);
                                }
                            }
                            catch { } // Ignore errors when closing

                            // Try to replace using the new path
                            occ.Replace(newPath, false);

                            // Force update after replacement
                            ((Inventor.Document)asmDoc).Update();
                            Thread.Sleep(1000);

                            return true;
                        }
                    }
                    catch (Exception ex1)
                    {
                        Console.WriteLine($"Method 1 - Replace(path, false) failed: {ex1.Message}");

                        // Method 2: Try with different parameters
                        try
                        {
                            // Force update again before second attempt
                            ((Inventor.Document)asmDoc).Update();
                            Thread.Sleep(1000);

                            // Get the assembly definition
                            var asmDef = (AssemblyComponentDefinition)occ.ContextDefinition;

                            // Try to create the new occurrence first
                            var newOcc = asmDef.Occurrences.Add(newPath, transform);

                            // Apply properties
                            newOcc.Visible = visible;
                            if (isSuppressed)
                            {
                                newOcc.Suppress();
                            }

                            // Try to preserve the name if possible
                            try
                            {
                                if (!string.IsNullOrEmpty(name) && name != newOcc.Name)
                                {
                                    newOcc.Name = name;
                                }
                            }
                            catch { } // Ignore name setting errors

                            // Force update before deletion
                            ((Inventor.Document)asmDoc).Update();
                            Thread.Sleep(1000);

                            // Delete the old occurrence
                            occ.Delete();

                            // Final update
                            ((Inventor.Document)asmDoc).Update();

                            return true;
                        }
                        catch (Exception ex2)
                        {
                            Console.WriteLine($"Method 2 - Delete and recreate failed: {ex2.Message}");

                            // Method 3: Try using the document's reference update capabilities
                            try
                            {
                                // Force update before final attempt
                                ((Inventor.Document)asmDoc).Update();
                                Thread.Sleep(1000);

                                // Try to update the reference at the document level
                                var refDocuments = asmDoc.ReferencedDocuments;
                                foreach (Document refDoc in refDocuments)
                                {
                                    if (string.Equals(refDoc.FullFileName, oldPath, StringComparison.OrdinalIgnoreCase))
                                    {
                                        // Try to close and reopen the reference
                                        ((Inventor.Document)refDoc).Close(false);
                                        Thread.Sleep(1000);

                                        // Force the assembly to update its references
                                        ((Inventor.Document)asmDoc).Update();
                                        return true;
                                    }
                                }
                            }
                            catch (Exception ex3)
                            {
                                Console.WriteLine($"Method 3 - Document reference update failed: {ex3.Message}");

                                if (attempt < maxRetries)
                                {
                                    Console.WriteLine($"Retrying in {retryDelayMs}ms...");
                                    Thread.Sleep(retryDelayMs);
                                    continue;
                                }
                                return false;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Critical error in TryReplaceReference (attempt {attempt}/{maxRetries}): {ex.Message}");
                    if (attempt < maxRetries)
                    {
                        Thread.Sleep(retryDelayMs);
                        continue;
                    }
                    return false;
                }
            }
            return false;
        }

        private bool RecreateOccurrenceImproved(ComponentOccurrence originalOcc, string newPath)
        {
            try
            {
                // Store all properties we need to preserve
                var transformation = originalOcc.Transformation;
                var visible = originalOcc.Visible;
                var name = originalOcc.Name;
                var isSuppressed = originalOcc.Suppressed;

                // Get the assembly definition
                var asmDef = (AssemblyComponentDefinition)originalOcc.ContextDefinition;

                // Store the index for placement
                int originalIndex = -1;
                for (int i = 1; i <= asmDef.Occurrences.Count; i++)
                {
                    if (asmDef.Occurrences[i] == originalOcc)
                    {
                        originalIndex = i;
                        break;
                    }
                }

                // Force update before recreation
                ((Inventor.Document)asmDef.Document).Update();
                Thread.Sleep(1000);

                // Create the new occurrence first
                var newOcc = asmDef.Occurrences.Add(newPath, transformation);

                // Apply properties
                newOcc.Visible = visible;

                // Handle suppression state
                if (isSuppressed)
                {
                    try
                    {
                        newOcc.Suppress();
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not suppress new occurrence: {ex.Message}");
                    }
                }

                // Try to preserve the name if possible
                try
                {
                    if (!string.IsNullOrEmpty(name) && name != newOcc.Name)
                    {
                        newOcc.Name = name;
                    }
                }
                catch
                {
                    // Name might not be settable, ignore
                }

                // Force update before deletion
                ((Inventor.Document)asmDef.Document).Update();
                Thread.Sleep(1000);

                // Delete the original occurrence
                originalOcc.Delete();

                // Final update
                ((Inventor.Document)asmDef.Document).Update();

                Console.WriteLine($"Successfully recreated occurrence: {System.IO.Path.GetFileName(newPath)}");
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"RecreateOccurrenceImproved failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Try to update references at the assembly document level
        /// </summary>
        private bool UpdateReferenceViaAssemblyDocument(AssemblyDocument asmDoc, string oldPath, string newPath)
        {
            try
            {
                // Try using the document's reference update capabilities
                var refDocuments = asmDoc.ReferencedDocuments;

                foreach (Document refDoc in refDocuments)
                {
                    if (string.Equals(refDoc.FullFileName, oldPath, StringComparison.OrdinalIgnoreCase))
                    {
                        // Try to close and reopen the reference
                        try
                        {
                            refDoc.Close(false);

                            // Force the assembly to update its references
                            ((Inventor.Document)asmDoc).Update();

                            return true;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Reference document update failed: {ex.Message}");
                            return false;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"UpdateReferenceViaAssemblyDocument failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Enhanced file lock checking with retry mechanism
        /// </summary>
        private bool IsFileLocked(string filePath, int maxRetries = 3)
        {
            for (int attempt = 1; attempt <= maxRetries; attempt++)
            {
                try
                {
                    using (var stream = System.IO.File.Open(filePath, System.IO.FileMode.Open, System.IO.FileAccess.ReadWrite, System.IO.FileShare.None))
                    {
                        return false; // File is not locked
                    }
                }
                catch (System.IO.IOException)
                {
                    if (attempt < maxRetries)
                    {
                        Console.WriteLine($"File locked, attempt {attempt}/{maxRetries}: {filePath}");
                        Thread.Sleep(2000); // Wait 2 seconds before retry
                        continue;
                    }
                    return true; // File is locked after all retries
                }
                catch (Exception)
                {
                    return false; // If we can't check, assume it's not locked
                }
            }

            return true;
        }

        /// <summary>
        /// Enhanced Inventor cleanup with more aggressive COM object release
        /// </summary>
        private void ForceInventorCleanup(Inventor.Application inventorApp)
        {
            try
            {
                Console.WriteLine("Starting enhanced Inventor cleanup...");

                // Close all open documents
                foreach (Document doc in inventorApp.Documents)
                {
                    try
                    {
                        doc.Close(false);
                        Marshal.ReleaseComObject(doc);
                    }
                    catch { } // Ignore errors during cleanup
                }

                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                // Additional COM object cleanup
                try
                {
                    // Release any remaining COM objects
                    Marshal.FinalReleaseComObject(inventorApp);
                }
                catch { } // Ignore errors during final cleanup

                // Final garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                Console.WriteLine("Enhanced Inventor cleanup completed");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during Inventor cleanup: {ex.Message}");
            }
        }

        /// <summary>
        /// Wait for file system to stabilize after file operations
        /// </summary>
        private void WaitForFileSystemStabilization(int timeoutSeconds = 30)
        {
            Console.WriteLine($"Waiting for file system stabilization ({timeoutSeconds}s)...");

            var startTime = DateTime.Now;

            while ((DateTime.Now - startTime).TotalSeconds < timeoutSeconds)
            {
                // Force multiple garbage collections
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                Thread.Sleep(2000);

                // You could add additional checks here if needed
                // For example, checking if specific files are still locked
            }

            Console.WriteLine("File system stabilization wait completed");
        }
        /// <summary>
        /// Determines if a file is a Content Center file that should not be renamed
        /// </summary>
        private bool IsContentCenterFile(string filePath)
        {
            // Check if the path contains Content Center indicators
            return filePath.Contains("Content Center Files", StringComparison.OrdinalIgnoreCase) ||
                   filePath.Contains("Parker", StringComparison.OrdinalIgnoreCase) ||
                   filePath.Contains("ISO", StringComparison.OrdinalIgnoreCase) ||
                   filePath.Contains("ANSI", StringComparison.OrdinalIgnoreCase) ||
                   filePath.Contains("DIN", StringComparison.OrdinalIgnoreCase) ||
                   filePath.Contains("JIS", StringComparison.OrdinalIgnoreCase) ||
                   filePath.Contains("GB", StringComparison.OrdinalIgnoreCase);
        }

        /// <summary>
        /// Determines if a file should be renamed based on part number prefix matching from iProperties
        /// </summary>
        private bool ShouldRename(ComponentOccurrence occurrence, string partPrefix)
        {
            try
            {
                // Get the part number directly from the occurrence's referenced document
                Document referencedDoc = (Document)occurrence.ReferencedDocumentDescriptor.ReferencedDocument;

                if (referencedDoc == null)
                    return false;

                string partNumber = "";

                // Get the part number from iProperties
                if (referencedDoc is PartDocument partDoc)
                {
                    partNumber = partDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value?.ToString() ?? "";
                }
                else if (referencedDoc is AssemblyDocument asmDoc)
                {
                    partNumber = asmDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value?.ToString() ?? "";
                }

                if (string.IsNullOrWhiteSpace(partNumber))
                    return false;

                // Extract the first part before underscore or dash from part number
                string[] parts = partNumber.Split(new char[] { '_', '-' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length == 0)
                    return false;

                string firstPart = parts[0].Trim();

                // Check if the first part matches the part prefix (e.g., "ABC" matches "ABC")
                return string.Equals(firstPart, partPrefix, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading part number from occurrence: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Determines if a main assembly should be renamed based on part number prefix matching from iProperties
        /// </summary>
        private bool ShouldRenameAssembly(AssemblyDocument asmDoc, string partPrefix)
        {
            try
            {
                string partNumber = asmDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value?.ToString() ?? "";

                if (string.IsNullOrWhiteSpace(partNumber))
                    return false;

                // Extract the first part before underscore or dash from part number
                string[] parts = partNumber.Split(new char[] { '_', '-' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length == 0)
                    return false;

                string firstPart = parts[0].Trim();

                // Check if the first part matches the part prefix (e.g., "ABC" matches "ABC")
                return string.Equals(firstPart, partPrefix, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading part number from assembly: {ex.Message}");
                return false;
            }
        }
        private string GenerateUniqueFileName(string originalName, string newPrefix, string extension, string directory, HashSet<string> usedNames)
        {
            // Remove any existing prefix pattern (like ABC099001, XYZ1, etc.)
            string cleanName = originalName;

            // Remove common prefixes (adjust regex pattern as needed)
            var prefixPattern = @"^[A-Z]{2,}\d*_?";
            var match = System.Text.RegularExpressions.Regex.Match(cleanName, prefixPattern);
            if (match.Success)
            {
                cleanName = cleanName.Substring(match.Length);
            }

            // Ensure we have something to work with
            if (string.IsNullOrEmpty(cleanName))
            {
                cleanName = "Part1";
            }

            // Generate base name
            string baseName = $"{newPrefix}_{cleanName.TrimStart('_')}";
            string fullName = baseName + extension;

            // Check for conflicts and resolve
            int counter = 1;
            while (usedNames.Contains(fullName) || System.IO.File.Exists(System.IO.Path.Combine(directory, fullName)))
            {
                fullName = $"{baseName}_{counter:D2}{extension}";
                counter++;
            }

            return fullName;
        }
        private List<string> DiscoverAssemblyFiles(string drawingsPath)
        {
            try
            {
                // Get ALL .iam files in the directory - no filtering needed
                return Directory.GetFiles(drawingsPath, "*.iam", SearchOption.TopDirectoryOnly)
                    .Select(System.IO.Path.GetFileName)
                    .Where(name => !string.IsNullOrEmpty(name)) // Filter out any null/empty names
                    .OrderByDescending(name => name) // Descending order by name
                    .ToList()!; // Safe to use ! here since we filtered nulls above
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error discovering assembly files: {ex.Message}");
                return new List<string>();
            }
        }

        private void CleanupInventorApp()
        {
            if (_inventorApp != null)
            {
                _inventorApp.SilentOperation = false;

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

        /// <summary>
        /// Recursively renames assemblies and parts, updates references and properties, and returns a list of old file paths to delete.
        /// </summary>
        public List<string> RenameAssemblyRecursively(List<string> assemblyDocumentNames, Dictionary<string, string> fileNames)
        {
            var pathToDelete = new List<string>();
            var inventorApp = GetInventorApplication();
            inventorApp.SilentOperation = true;
            inventorApp.Visible = false;

            Console.WriteLine($"Starting recursive rename for {assemblyDocumentNames.Count} assemblies with {fileNames.Count} file mappings");

            foreach (var assemblyDocumentName in assemblyDocumentNames)
            {
                string assemblyFilePath = System.IO.Path.GetFullPath(assemblyDocumentName);
                if (!System.IO.File.Exists(assemblyFilePath))
                {
                    Console.WriteLine($"Assembly file not found: {assemblyFilePath}");
                    continue;
                }

                Console.WriteLine($"Processing assembly: {System.IO.Path.GetFileName(assemblyFilePath)}");
                AssemblyDocument? asmDoc = null;
                try
                {
                    asmDoc = (AssemblyDocument)inventorApp.Documents.Open(assemblyFilePath, true);
                    var docDescriptors = asmDoc.ReferencedDocumentDescriptors;
                    Console.WriteLine($"Found {docDescriptors.Count} referenced documents");

                    foreach (DocumentDescriptor oDocDescriptor in docDescriptors)
                    {
                        try
                        {
                            string referencedPath = oDocDescriptor.FullDocumentName;
                            if (!fileNames.TryGetValue(referencedPath, out var newFileName) || string.IsNullOrEmpty(newFileName))
                                continue;

                            string newFullName = System.IO.Path.GetFullPath(newFileName);
                            Console.WriteLine($"Processing: {System.IO.Path.GetFileName(referencedPath)} -> {System.IO.Path.GetFileName(newFullName)}");

                            if (!System.IO.File.Exists(newFullName))
                            {
                                // Get the referenced document to determine its type
                                Document? referencedDoc = null;
                                try
                                {
                                    referencedDoc = (Document)oDocDescriptor.ReferencedDocument;
                                    if (referencedDoc == null)
                                        continue;

                                    if (referencedDoc is PartDocument)
                                    {
                                        PartDocument? partDoc = null;
                                        try
                                        {
                                            Console.WriteLine($"Opening part document: {System.IO.Path.GetFileName(referencedPath)}");
                                            partDoc = (PartDocument)inventorApp.Documents.Open(referencedPath, true);
                                            pathToDelete.Add(referencedPath);
                                            
                                            Console.WriteLine($"Saving part as: {System.IO.Path.GetFileName(newFullName)}");
                                            partDoc.SaveAs(newFullName, false);
                                            
                                            // Update part number property
                                            try
                                            {
                                                var designProps = partDoc.PropertySets["Design Tracking Properties"];
                                                designProps["Part Number"].Value = System.IO.Path.GetFileNameWithoutExtension(newFullName);
                                                Console.WriteLine($"Updated part number to: {System.IO.Path.GetFileNameWithoutExtension(newFullName)}");
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"Warning: Could not update part number: {ex.Message}");
                                            }
                                            
                                            // Change material if Generic
                                            try
                                            {
                                                var mat = partDoc.ComponentDefinition.Material;
                                                if (mat.Name == "Generic")
                                                {
                                                    // Optionally set to a default material, e.g., "Steel"
                                                    partDoc.ComponentDefinition.Material = partDoc.Materials["Steel"];
                                                    Console.WriteLine("Changed material from Generic to Steel");
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"Warning: Could not change material: {ex.Message}");
                                            }
                                            
                                            partDoc.Update();
                                            Console.WriteLine("Part document updated successfully");
                                            
                                            // Replace reference in parent
                                            try
                                            {
                                                referencedDoc.Save();
                                                Console.WriteLine("Referenced document saved successfully");
                                            }
                                            catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                                            {
                                                Console.WriteLine($"Warning: Could not save referenced document (E_FAIL): {comEx.Message}");
                                                Console.WriteLine($"File: {referencedPath}");
                                                // Continue processing - don't fail the entire operation
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"Warning: Could not save referenced document: {ex.Message}");
                                                Console.WriteLine($"File: {referencedPath}");
                                                // Continue processing - don't fail the entire operation
                                            }
                                        }
                                        catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                                        {
                                            Console.WriteLine($"COM Error (E_FAIL) processing part document: {comEx.Message}");
                                            Console.WriteLine($"File: {referencedPath}");
                                            throw;
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Error processing part document: {ex.Message}");
                                            Console.WriteLine($"File: {referencedPath}");
                                            throw;
                                        }
                                        finally
                                        {
                                            if (partDoc != null)
                                            {
                                                try
                                                {
                                                    partDoc.Close(true);
                                                    Marshal.ReleaseComObject(partDoc);
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine($"Error closing part document: {ex.Message}");
                                                }
                                            }
                                        }
                                    }
                                    else if (referencedDoc is AssemblyDocument)
                                    {
                                        AssemblyDocument? subAsmDoc = null;
                                        try
                                        {
                                            Console.WriteLine($"Opening subassembly document: {System.IO.Path.GetFileName(referencedPath)}");
                                            subAsmDoc = (AssemblyDocument)inventorApp.Documents.Open(referencedPath, true);
                                            pathToDelete.Add(referencedPath);
                                            
                                            Console.WriteLine($"Saving subassembly as: {System.IO.Path.GetFileName(newFullName)}");
                                            subAsmDoc.SaveAs(newFullName, false);
                                            
                                            // Update assembly number property
                                            try
                                            {
                                                var designProps = subAsmDoc.PropertySets["Design Tracking Properties"];
                                                designProps["Part Number"].Value = System.IO.Path.GetFileNameWithoutExtension(newFullName);
                                                Console.WriteLine($"Updated assembly number to: {System.IO.Path.GetFileNameWithoutExtension(newFullName)}");
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"Warning: Could not update assembly number: {ex.Message}");
                                            }
                                            
                                            // Enable & sort BOM (if needed)
                                            try
                                            {
                                                var bom = subAsmDoc.ComponentDefinition.BOM;
                                                bom.StructuredViewEnabled = true;
                                                bom.StructuredViewFirstLevelOnly = false;
                                                // Note: StructuredViewSortColumn might not be available in all Inventor versions
                                                // bom.StructuredViewSortColumn = "Part Number";
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"Warning: Could not configure BOM: {ex.Message}");
                                            }
                                            
                                            subAsmDoc.Update();
                                            Console.WriteLine("Subassembly document updated successfully");
                                            
                                            // Replace reference in parent
                                            try
                                            {
                                                referencedDoc.Save();
                                                Console.WriteLine("Referenced document saved successfully");
                                            }
                                            catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                                            {
                                                Console.WriteLine($"Warning: Could not save referenced document (E_FAIL): {comEx.Message}");
                                                Console.WriteLine($"File: {referencedPath}");
                                                // Continue processing - don't fail the entire operation
                                            }
                                            catch (Exception ex)
                                            {
                                                Console.WriteLine($"Warning: Could not save referenced document: {ex.Message}");
                                                Console.WriteLine($"File: {referencedPath}");
                                                // Continue processing - don't fail the entire operation
                                            }
                                        }
                                        catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                                        {
                                            Console.WriteLine($"COM Error (E_FAIL) processing subassembly document: {comEx.Message}");
                                            Console.WriteLine($"File: {referencedPath}");
                                            throw;
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Error processing subassembly document: {ex.Message}");
                                            Console.WriteLine($"File: {referencedPath}");
                                            throw;
                                        }
                                        finally
                                        {
                                            if (subAsmDoc != null)
                                            {
                                                try
                                                {
                                                    subAsmDoc.Close(true);
                                                    Marshal.ReleaseComObject(subAsmDoc);
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine($"Error closing subassembly document: {ex.Message}");
                                                }
                                            }
                                        }
                                    }
                                }
                                catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                                {
                                    Console.WriteLine($"COM Error (E_FAIL) accessing referenced document: {comEx.Message}");
                                    Console.WriteLine($"File: {referencedPath}");
                                    // Continue processing - don't fail the entire operation
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error accessing referenced document: {ex.Message}");
                                    Console.WriteLine($"File: {referencedPath}");
                                    // Continue processing - don't fail the entire operation
                                }
                                finally
                                {
                                    if (referencedDoc != null)
                                    {
                                        Marshal.ReleaseComObject(referencedDoc);
                                    }
                                }
                            }
                            // Note: Reference replacement is now handled after the file processing block

                            // CRITICAL FIX: Always replace the reference in the parent assembly, regardless of whether the file existed or not
                            // This ensures that the component names in the browser tree are updated
                            try
                            {
                                Console.WriteLine($"Updating reference in parent assembly: {System.IO.Path.GetFileName(referencedPath)} -> {System.IO.Path.GetFileName(newFullName)}");
                                
                                // Find and replace the occurrence in the parent assembly
                                bool referenceUpdated = false;
                                foreach (ComponentOccurrence occ in asmDoc.ComponentDefinition.Occurrences)
                                {
                                    try
                                    {
                                        string occRefPath = occ.ReferencedDocumentDescriptor.FullDocumentName;
                                        if (occRefPath.Equals(referencedPath, StringComparison.OrdinalIgnoreCase))
                                        {
                                            Console.WriteLine($"Found matching occurrence: {occ.Name}, replacing reference...");
                                            occ.Replace(newFullName, false);
                                            referenceUpdated = true;
                                            Console.WriteLine($"Successfully replaced reference for occurrence: {occ.Name}");
                                            break;
                                        }
                                    }
                                    catch (Exception occEx)
                                    {
                                        Console.WriteLine($"Warning: Error checking occurrence {occ.Name}: {occEx.Message}");
                                        continue;
                                    }
                                }

                                if (!referenceUpdated)
                                {
                                    Console.WriteLine($"Warning: Could not find occurrence to replace for: {System.IO.Path.GetFileName(referencedPath)}");
                                }
                            }
                            catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                            {
                                Console.WriteLine($"Warning: E_FAIL error updating reference: {System.IO.Path.GetFileName(referencedPath)} -> {System.IO.Path.GetFileName(newFullName)}: {comEx.Message}");
                                // Continue processing - don't fail the entire operation
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error updating reference: {System.IO.Path.GetFileName(referencedPath)} -> {System.IO.Path.GetFileName(newFullName)}: {ex.Message}");
                                // Continue processing - don't fail the entire operation
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error processing document descriptor: {ex.Message}");
                            // Continue processing - don't fail the entire operation
                        }
                    }
                    
                    // After all descriptors processed
                    try
                    {
                        var bom = asmDoc.ComponentDefinition.BOM;
                        bom.StructuredViewEnabled = true;
                        // Note: StructuredViewSortColumn might not be available in all Inventor versions
                        // bom.StructuredViewSortColumn = "Part Number";
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not configure BOM: {ex.Message}");
                    }
                    
                    // Force update and save the assembly to ensure all reference changes are applied
                    try
                    {
                        Console.WriteLine($"Forcing update and save of assembly: {System.IO.Path.GetFileName(assemblyFilePath)}");
                        asmDoc.Update();
                        Thread.Sleep(1000); // Give Inventor time to process the update
                        asmDoc.Save2(true); // Save with Yes to All, suppress dialogs
                        Thread.Sleep(1000); // Give Inventor time to save
                        Console.WriteLine($"Assembly updated and saved successfully: {System.IO.Path.GetFileName(assemblyFilePath)}");
                    }
                    catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                    {
                        Console.WriteLine($"Warning: E_FAIL error saving assembly: {comEx.Message}");
                        // Try to save without the update
                        try
                        {
                            asmDoc.Save2(true);
                            Console.WriteLine($"Assembly saved successfully (without update): {System.IO.Path.GetFileName(assemblyFilePath)}");
                        }
                        catch (Exception saveEx)
                        {
                            Console.WriteLine($"Error saving assembly: {saveEx.Message}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error updating/saving assembly: {ex.Message}");
                    }
                }
                catch (System.Runtime.InteropServices.COMException comEx) when ((uint)comEx.ErrorCode == 0x80004005)
                {
                    Console.WriteLine($"COM Error (E_FAIL) processing assembly: {comEx.Message}");
                    Console.WriteLine($"File: {assemblyFilePath}");
                    // Continue processing - don't fail the entire operation
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing assembly: {ex.Message}");
                    Console.WriteLine($"File: {assemblyFilePath}");
                    // Continue processing - don't fail the entire operation
                }
                finally
                {
                    if (asmDoc != null)
                    {
                        try
                        {
                            asmDoc.Close(true);
                            Marshal.ReleaseComObject(asmDoc);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error closing assembly document: {ex.Message}");
                        }
                    }
                }
            }
            return pathToDelete;
        }

        /// <summary>
        /// Recursively renames assemblies and parts using a prefix, automatically discovering files and generating rename mappings.
        /// </summary>
        public List<string> RenameAssemblyRecursivelyWithPrefix(string modelPath, string prefix)
        {
            var pathToDelete = new List<string>();
            var inventorApp = GetInventorApplication();
            inventorApp.SilentOperation = true;
            inventorApp.Visible = false;

            try
            {
                Console.WriteLine($"=== Starting Recursive Rename with Prefix ===");
                Console.WriteLine($"Model Path: {modelPath}");
                Console.WriteLine($"Prefix: {prefix}");

                // Discover all assembly files in the model path
                var assemblyFiles = DiscoverAssemblyFiles(modelPath);
                if (assemblyFiles.Count == 0)
                {
                    Console.WriteLine("No assembly files found in the specified path.");
                    return pathToDelete;
                }

                Console.WriteLine($"Found {assemblyFiles.Count} assembly files to process.");

                // Build rename mappings based on part numbers and prefix
                var fileNames = new Dictionary<string, string>();
                var usedNames = new HashSet<string>();

                foreach (var assemblyFile in assemblyFiles)
                {
                    string assemblyFilePath = System.IO.Path.Combine(modelPath, assemblyFile);
                    if (!System.IO.File.Exists(assemblyFilePath))
                    {
                        Console.WriteLine($"Assembly file not found: {assemblyFilePath}");
                        continue;
                    }

                    Console.WriteLine($"Processing assembly: {assemblyFile}");
                    AssemblyDocument? asmDoc = null;
                    try
                    {
                        asmDoc = (AssemblyDocument)inventorApp.Documents.Open(assemblyFilePath, true);
                        var docDescriptors = asmDoc.ReferencedDocumentDescriptors;
                        Console.WriteLine($"Found {docDescriptors.Count} referenced documents in {assemblyFile}");

                        foreach (DocumentDescriptor oDocDescriptor in docDescriptors)
                        {
                            try
                            {
                                string referencedPath = oDocDescriptor.FullDocumentName;
                                if (IsContentCenterFile(referencedPath))
                                {
                                    Console.WriteLine($"Skipping Content Center file: {System.IO.Path.GetFileName(referencedPath)}");
                                    continue;
                                }

                                string fileName = System.IO.Path.GetFileNameWithoutExtension(referencedPath);
                                string fileExt = System.IO.Path.GetExtension(referencedPath);
                                string dir = System.IO.Path.GetDirectoryName(referencedPath)!;

                                // Check if it should be renamed based on part number
                                if (!ShouldRenameByPartNumber(oDocDescriptor, prefix))
                                {
                                    Console.WriteLine($"Skipping file (doesn't need rename): {fileName}");
                                    continue;
                                }

                                // Skip if already starts with new prefix
                                if (fileName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                                {
                                    Console.WriteLine($"Skipping file (already has correct prefix): {fileName}");
                                    continue;
                                }

                                // Generate new name with conflict resolution
                                string newFileName = GenerateUniqueFileName(fileName, prefix, fileExt, dir, usedNames);
                                string newPath = System.IO.Path.Combine(dir, newFileName);

                                if (!fileNames.ContainsKey(referencedPath))
                                {
                                    fileNames.Add(referencedPath, newPath);
                                    usedNames.Add(newFileName);
                                    Console.WriteLine($"Added to rename list: {fileName} -> {newFileName}");
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error processing document descriptor: {ex.Message}");
                            }
                        }

                        // Handle main assembly renaming
                        string mainFileName = System.IO.Path.GetFileNameWithoutExtension(assemblyFilePath);
                        if (ShouldRenameAssemblyByPartNumber(asmDoc, prefix) && !mainFileName.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                        {
                            string mainExt = System.IO.Path.GetExtension(assemblyFilePath);
                            string newMainFileName = GenerateUniqueFileName(mainFileName, prefix, mainExt, modelPath, usedNames);
                            string mainNewPath = System.IO.Path.Combine(modelPath, newMainFileName);

                            if (!fileNames.ContainsKey(assemblyFilePath))
                            {
                                fileNames.Add(assemblyFilePath, mainNewPath);
                                usedNames.Add(newMainFileName);
                                Console.WriteLine($"Added main assembly to rename list: {mainFileName} -> {newMainFileName}");
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing assembly {assemblyFile}: {ex.Message}");
                    }
                    finally
                    {
                        if (asmDoc != null)
                        {
                            try
                            {
                                asmDoc.Close(false);
                                Marshal.ReleaseComObject(asmDoc);
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine($"Error closing assembly document: {ex.Message}");
                            }
                        }
                    }
                }

                if (fileNames.Count == 0)
                {
                    Console.WriteLine("No files found that need renaming.");
                    return pathToDelete;
                }

                Console.WriteLine($"Found {fileNames.Count} files to rename.");

                // Call the recursive rename method
                try
                {
                    Console.WriteLine("=== Starting Recursive Rename Operation ===");
                    pathToDelete = RenameAssemblyRecursively(assemblyFiles.Select(f => System.IO.Path.Combine(modelPath, f)).ToList(), fileNames);
                    Console.WriteLine($"Recursive rename completed. {pathToDelete.Count} files can be deleted.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during recursive rename operation: {ex.Message}");
                    Console.WriteLine($"Stack trace: {ex.StackTrace}");
                    throw;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during recursive rename with prefix: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                throw;
            }
            finally
            {
                try
                {
                    CleanupInventorApp();
                    GC.Collect();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during cleanup: {ex.Message}");
                }
            }

            return pathToDelete;
        }

        /// <summary>
        /// Determines if a file should be renamed based on part number prefix matching from the document descriptor.
        /// </summary>
        private bool ShouldRenameByPartNumber(DocumentDescriptor docDescriptor, string partPrefix)
        {
            try
            {
                Document? referencedDoc = (Document)docDescriptor.ReferencedDocument;
                if (referencedDoc == null)
                    return false;

                string partNumber = "";

                // Get the part number from iProperties
                if (referencedDoc is PartDocument partDoc)
                {
                    partNumber = partDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value?.ToString() ?? "";
                }
                else if (referencedDoc is AssemblyDocument asmDoc)
                {
                    partNumber = asmDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value?.ToString() ?? "";
                }

                if (string.IsNullOrWhiteSpace(partNumber))
                    return false;

                // Extract the first part before underscore or dash from part number
                string[] parts = partNumber.Split(new char[] { '_', '-' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length == 0)
                    return false;

                string firstPart = parts[0].Trim();

                // Check if the first part matches the part prefix
                return string.Equals(firstPart, partPrefix, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading part number from document descriptor: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Determines if a main assembly should be renamed based on part number prefix matching.
        /// </summary>
        private bool ShouldRenameAssemblyByPartNumber(AssemblyDocument asmDoc, string partPrefix)
        {
            try
            {
                string partNumber = asmDoc.PropertySets["Design Tracking Properties"]["Part Number"].Value?.ToString() ?? "";

                if (string.IsNullOrWhiteSpace(partNumber))
                    return false;

                // Extract the first part before underscore or dash from part number
                string[] parts = partNumber.Split(new char[] { '_', '-' }, StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length == 0)
                    return false;

                string firstPart = parts[0].Trim();

                // Check if the first part matches the part prefix
                return string.Equals(firstPart, partPrefix, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading part number from assembly: {ex.Message}");
                return false;
            }
        }
    }
}

