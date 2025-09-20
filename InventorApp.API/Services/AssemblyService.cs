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

                    // Configure Inventor for automation to handle dialogs automatically
                    ConfigureInventorForAutomation(_inventorApp);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to initialize Inventor application: {ex.Message}. Please ensure Inventor is running and properly registered.", ex);
                }
            }
            return _inventorApp;
        }

        /// <summary>
        /// Configures Inventor application for silent automation to prevent user dialogs.
        /// </summary>
        private void ConfigureInventorForAutomation(Inventor.Application inventorApp)
        {
            try
            {
                // Enable silent operation to prevent user dialogs
                inventorApp.SilentOperation = true;

                // Set user interaction level to suppress dialogs
                inventorApp.UserInterfaceManager.UserInteractionDisabled = true;

                Console.WriteLine("Inventor configured for silent automation");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Warning: Could not configure Inventor for automation: {ex.Message}");
                // Continue with basic silent operation setting
                try
                {
                    inventorApp.SilentOperation = true;
                    inventorApp.UserInterfaceManager.UserInteractionDisabled = true;
                }
                catch (Exception silentEx)
                {
                    Console.WriteLine($"Warning: Could not set SilentOperation: {silentEx.Message}");
                }
            }
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
            dynamic? inventorApp = null;
            dynamic? doc = null;
            try
            {
                Type? inventorType = Type.GetTypeFromProgID("Inventor.Application");
                if (inventorType == null)
                    throw new Exception("Inventor is not installed.");
                inventorApp = Activator.CreateInstance(inventorType);
#pragma warning disable CS8602 // Dereference of a possibly null reference.
                inventorApp.Visible = false;
#pragma warning restore CS8602 // Dereference of a possibly null reference.
                doc = inventorApp.Documents.Open(filePath, true);
                string partNumber = "";

                var propSets = doc.PropertySets;
                var designProps = propSets["Design Tracking Properties"];
                partNumber = designProps["Part Number"].Value.ToString();

                return partNumber ?? "";
            }
            catch
            {
                return "";
            }
            finally
            {
                try
                {
                    if (doc != null)
                    {
                        doc.Close();
                        Marshal.ReleaseComObject(doc);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error closing document in GetPartNumberFromFile: {ex.Message}");
                }

                try
                {
                    if (inventorApp != null)
                    {
                        inventorApp.Quit();
                        Marshal.ReleaseComObject(inventorApp);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error quitting Inventor in GetPartNumberFromFile: {ex.Message}");
                }
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

            // First, handle main assemblies that need to be renamed
            var mainAssembliesToRename = new List<string>();
            foreach (var assemblyDocumentName in assemblyDocumentNames)
            {
                string assemblyFilePath = System.IO.Path.GetFullPath(assemblyDocumentName);
                if (fileNames.ContainsKey(assemblyFilePath))
                {
                    mainAssembliesToRename.Add(assemblyFilePath);
                }
            }

            // Process main assemblies first
            foreach (var assemblyFilePath in mainAssembliesToRename)
            {
                if (!System.IO.File.Exists(assemblyFilePath))
                {
                    Console.WriteLine($"Main assembly file not found: {assemblyFilePath}");
                    continue;
                }

                string newAssemblyPath = fileNames[assemblyFilePath];
                Console.WriteLine($"Processing main assembly: {System.IO.Path.GetFileName(assemblyFilePath)} -> {System.IO.Path.GetFileName(newAssemblyPath)}");

                AssemblyDocument? asmDoc = null;
                try
                {
                    asmDoc = (AssemblyDocument)inventorApp.Documents.Open(assemblyFilePath, true);
                    pathToDelete.Add(assemblyFilePath);

                    Console.WriteLine($"Saving main assembly as: {System.IO.Path.GetFileName(newAssemblyPath)}");
                    asmDoc.SaveAs(newAssemblyPath, false);

                    // Update assembly number property
                    try
                    {
                        var designProps = asmDoc.PropertySets["Design Tracking Properties"];
                        designProps["Part Number"].Value = System.IO.Path.GetFileNameWithoutExtension(newAssemblyPath);
                        Console.WriteLine($"Updated main assembly number to: {System.IO.Path.GetFileNameWithoutExtension(newAssemblyPath)}");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not update main assembly number: {ex.Message}");
                    }

                    // Enable & sort BOM (if needed)
                    try
                    {
                        var bom = asmDoc.ComponentDefinition.BOM;
                        bom.StructuredViewEnabled = true;
                        bom.StructuredViewFirstLevelOnly = false;
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Warning: Could not configure BOM: {ex.Message}");
                    }

                    asmDoc.Update();
                    Console.WriteLine("Main assembly document updated successfully");
                    asmDoc.Save();
                    Console.WriteLine("Main assembly saved successfully");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error processing main assembly {assemblyFilePath}: {ex.Message}");
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
                            Console.WriteLine($"Error closing main assembly document: {ex.Message}");
                        }
                    }
                }
            }

            // Now process all assemblies for their referenced documents
            foreach (var assemblyDocumentName in assemblyDocumentNames)
            {
                string assemblyFilePath = System.IO.Path.GetFullPath(assemblyDocumentName);

                // Check if this assembly was renamed and use the new path
                string currentAssemblyPath = assemblyFilePath;
                if (fileNames.ContainsKey(assemblyFilePath))
                {
                    currentAssemblyPath = fileNames[assemblyFilePath];
                }

                if (!System.IO.File.Exists(currentAssemblyPath))
                {
                    Console.WriteLine($"Assembly file not found: {currentAssemblyPath}");
                    continue;
                }

                Console.WriteLine($"Processing assembly: {System.IO.Path.GetFileName(currentAssemblyPath)}");
                AssemblyDocument? asmDoc = null;
                try
                {
                    asmDoc = (AssemblyDocument)inventorApp.Documents.Open(currentAssemblyPath, true);
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

                                // Find and replace the occurrence in the parent assembly using multiple strategies
                                bool referenceUpdated = false;
                                string originalFileName = System.IO.Path.GetFileNameWithoutExtension(referencedPath);
                                string targetFileName = System.IO.Path.GetFileNameWithoutExtension(newFullName);

                                // Strategy 1: Try to find by exact path match
                                foreach (ComponentOccurrence occ in asmDoc.ComponentDefinition.Occurrences)
                                {
                                    try
                                    {
                                        string occRefPath = occ.ReferencedDocumentDescriptor.FullDocumentName;
                                        if (occRefPath.Equals(referencedPath, StringComparison.OrdinalIgnoreCase))
                                        {
                                            Console.WriteLine($"Found matching occurrence by path: {occ.Name}, replacing reference...");
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

                                // Strategy 2: If not found by path, try to find by component name
                                if (!referenceUpdated)
                                {
                                    foreach (ComponentOccurrence occ in asmDoc.ComponentDefinition.Occurrences)
                                    {
                                        try
                                        {
                                            // Check if the occurrence name matches the original file name
                                            if (occ.Name.Equals(originalFileName, StringComparison.OrdinalIgnoreCase) ||
                                                occ.Name.StartsWith(originalFileName + ":", StringComparison.OrdinalIgnoreCase))
                                            {
                                                Console.WriteLine($"Found matching occurrence by name: {occ.Name}, replacing reference...");
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
                                }

                                // Strategy 3: If still not found, try to find by partial name match
                                if (!referenceUpdated)
                                {
                                    foreach (ComponentOccurrence occ in asmDoc.ComponentDefinition.Occurrences)
                                    {
                                        try
                                        {
                                            // Check if the occurrence name contains the original file name
                                            if (occ.Name.Contains(originalFileName, StringComparison.OrdinalIgnoreCase))
                                            {
                                                Console.WriteLine($"Found matching occurrence by partial name: {occ.Name}, replacing reference...");
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
                                }

                                if (!referenceUpdated)
                                {
                                    Console.WriteLine($"Warning: Could not find occurrence to replace for: {System.IO.Path.GetFileName(referencedPath)}");
                                    Console.WriteLine($"Searched for: {originalFileName} -> {targetFileName} in {asmDoc.ComponentDefinition.Occurrences.Count} occurrences");

                                    // Log all occurrences for debugging
                                    Console.WriteLine("Available occurrences:");
                                    foreach (ComponentOccurrence occ in asmDoc.ComponentDefinition.Occurrences)
                                    {
                                        try
                                        {
                                            string occRefPath = occ.ReferencedDocumentDescriptor.FullDocumentName;
                                            Console.WriteLine($"  - {occ.Name} -> {System.IO.Path.GetFileName(occRefPath)}");
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"  - {occ.Name} -> Error getting path: {ex.Message}");
                                        }
                                    }
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
                        Console.WriteLine($"Forcing update and save of assembly: {System.IO.Path.GetFileName(currentAssemblyPath)}");
                        asmDoc.Update();
                        Thread.Sleep(1000); // Give Inventor time to process the update
                        asmDoc.Save2(true); // Save with Yes to All, suppress dialogs
                        Thread.Sleep(1000); // Give Inventor time to save
                        Console.WriteLine($"Assembly updated and saved successfully: {System.IO.Path.GetFileName(currentAssemblyPath)}");
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
                    Console.WriteLine("No assembly files found in the specifiedSystem.IO.Path.");
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
        /// Recursively renames assemblies and parts using a prefix and updates drawing references.
        /// </summary>
        public DrawingUpdateResult RenameAssemblyRecursivelyWithPrefixAndUpdateDrawings(
            string modelPath,
            string drawingsPath,
            string projectPath,
            string oldPrefix,
            string newPrefix)
        {
            var result = new DrawingUpdateResult();
            var inventorApp = GetInventorApplication();
            inventorApp.SilentOperation = true;
            inventorApp.Visible = true; // Keep visible for debugging, but configure for automation

            // Ensure automation settings are applied
            ConfigureInventorForAutomation(inventorApp);

            try
            {
                Console.WriteLine($"=== Starting Enhanced Recursive Rename with Drawing Updates ===");
                Console.WriteLine($"Model Path: {modelPath}");
                Console.WriteLine($"Drawings Path: {drawingsPath}");
                Console.WriteLine($"Project Path: {projectPath}");
                Console.WriteLine($"Old Prefix: {oldPrefix}");
                Console.WriteLine($"New Prefix: {newPrefix}");

                // Step 1: Perform the recursive rename on model files
                Console.WriteLine("=== Step 1: Renaming Model Files ===");
                var filesToDelete = RenameAssemblyRecursivelyWithPrefix(modelPath, newPrefix);
                result.FilesToDelete = filesToDelete ?? new List<string>();

                // Step 2: Discover drawing files
                Console.WriteLine("=== Step 2: Discovering Drawing Files ===");
                var drawingFiles = DiscoverDrawingFiles(drawingsPath);
                Console.WriteLine($"Found {drawingFiles.Count} drawing files to process.");

                if (drawingFiles.Count == 0)
                {
                    Console.WriteLine("No drawing files found. Skipping drawing updates.");
                    return result;
                }

                // Step 3: Build file mapping for reference updates
                Console.WriteLine("=== Step 3: Building File Mapping ===");
                var fileMapping = BuildFileMappingForDrawingUpdates(modelPath, oldPrefix, newPrefix);
                Console.WriteLine($"Built mapping for {fileMapping.Count} files.");

                // Step 4: Update drawing references
                Console.WriteLine("=== Step 4: Updating Drawing References ===");
                UpdateDrawingReferences(drawingFiles, fileMapping, result, oldPrefix, newPrefix);

                // Step 5: Close Inventor before file operations to prevent file locks
                Console.WriteLine("=== Step 5: Closing Inventor for File Operations ===");
                try
                {
                    CleanupInventorApp();
                    GC.Collect();
                    Thread.Sleep(2000); // Wait for file handles to be released
                    Console.WriteLine("Inventor closed successfully for file operations");
                }
                catch (Exception cleanupEx)
                {
                    Console.WriteLine($"Warning: Error during Inventor cleanup: {cleanupEx.Message}");
                }

                // Step 6: Rename drawing files
                Console.WriteLine("=== Step 6: Renaming Drawing Files ===");
                var renamedDrawingFiles = RenameDrawingFiles(drawingFiles, oldPrefix, newPrefix, result);

                // Step 7: Update project files if project path is provided
                if (!string.IsNullOrEmpty(projectPath) && Directory.Exists(projectPath))
                {
                    Console.WriteLine("=== Step 7: Updating Project Files ===");
                    UpdateProjectFiles(projectPath, oldPrefix, newPrefix, result);
                    
                    // Step 8: Rename project file
                    Console.WriteLine("=== Step 8: Renaming Project File ===");
                    RenameProjectFile(projectPath, oldPrefix, newPrefix, result);
                }

                // Step 9: Validate drawing links
                Console.WriteLine("=== Step 9: Validating Drawing Links ===");
                ValidateDrawingLinks(renamedDrawingFiles, fileMapping, result, oldPrefix, newPrefix);

                Console.WriteLine("=== Enhanced Recursive Rename with Drawing Updates Completed ===");
                return result;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during enhanced recursive rename with drawing updates: {ex.Message}");
                Console.WriteLine($"Stack trace: {ex.StackTrace}");
                result.ErrorMessage = ex.Message;
                throw;
            }
            finally
            {
                try
                {
                    // Inventor is already closed in Step 5, but ensure cleanup in case of errors
                    if (_inventorApp != null)
                    {
                        Console.WriteLine("Final cleanup: Ensuring Inventor is closed");
                        CleanupInventorApp();
                        GC.Collect();
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error during final cleanup: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Renames drawing files to match the new prefix.
        /// </summary>
        private List<string> RenameDrawingFiles(List<string> drawingFiles, string oldPrefix, string newPrefix, DrawingUpdateResult result)
        {
            var renamedFiles = new List<string>();
            
            try
            {
                Console.WriteLine($"Renaming {drawingFiles.Count} drawing files...");
                
                foreach (var drawingFile in drawingFiles)
                {
                    try
                    {
                        var fileName = System.IO.Path.GetFileNameWithoutExtension(drawingFile);
                        var fileExt = System.IO.Path.GetExtension(drawingFile);
                        var directory = System.IO.Path.GetDirectoryName(drawingFile);
                        
                        // Check if the file name starts with the old prefix
                        if (fileName.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            // Generate new file name with new prefix
                            var newFileName = fileName.Replace(oldPrefix, newPrefix, StringComparison.OrdinalIgnoreCase);
                            
                            if (directory != null)
                            {
                                var newFilePath = System.IO.Path.Combine(directory, newFileName + fileExt);
                                
                                // Check if the new file already exists
                                if (System.IO.File.Exists(newFilePath))
                                {
                                    Console.WriteLine($"  Skipping {System.IO.Path.GetFileName(drawingFile)} - target file already exists");
                                    renamedFiles.Add(drawingFile); // Keep original path
                                    continue;
                                }
                                
                                // Rename the file with retry logic for file locks
                                bool renameSuccess = false;
                                for (int retry = 0; retry < 3; retry++)
                                {
                                    try
                                    {
                                        System.IO.File.Move(drawingFile, newFilePath);
                                        renameSuccess = true;
                                        break;
                                    }
                                    catch (IOException ioEx) when (ioEx.Message.Contains("being used by another process"))
                                    {
                                        Console.WriteLine($"  File locked, retrying in 1 second... (attempt {retry + 1}/3)");
                                        Thread.Sleep(1000);
                                    }
                                }
                                
                                if (renameSuccess)
                                {
                                    Console.WriteLine($"  Renamed: {System.IO.Path.GetFileName(drawingFile)} -> {System.IO.Path.GetFileName(newFilePath)}");
                                    renamedFiles.Add(newFilePath);
                                }
                                else
                                {
                                    Console.WriteLine($"  Failed to rename {System.IO.Path.GetFileName(drawingFile)} - file still locked");
                                    result.FailedDrawings.Add($"Failed to rename {System.IO.Path.GetFileName(drawingFile)} - file locked");
                                    renamedFiles.Add(drawingFile); // Keep original path
                                }
                                
                                // Add to files to delete list (original file)
                                result.FilesToDelete.Add(drawingFile);
                            }
                            else
                            {
                                Console.WriteLine($"  Could not get directory for: {drawingFile}");
                                renamedFiles.Add(drawingFile);
                            }
                        }
                        else
                        {
                            Console.WriteLine($"  Skipping {System.IO.Path.GetFileName(drawingFile)} - does not start with old prefix");
                            renamedFiles.Add(drawingFile);
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Error renaming drawing file {System.IO.Path.GetFileName(drawingFile)}: {ex.Message}");
                        result.FailedDrawings.Add($"Failed to rename {System.IO.Path.GetFileName(drawingFile)}: {ex.Message}");
                        renamedFiles.Add(drawingFile); // Keep original path on error
                    }
                }
                
                Console.WriteLine($"Drawing file renaming completed. {renamedFiles.Count} files processed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during drawing file renaming: {ex.Message}");
                result.ErrorMessage = $"Drawing file renaming error: {ex.Message}";
            }
            
            return renamedFiles;
        }

        /// <summary>
        /// Renames the project file (.ipj) to match the new prefix.
        /// </summary>
        private void RenameProjectFile(string projectPath, string oldPrefix, string newPrefix, DrawingUpdateResult result)
        {
            try
            {
                Console.WriteLine("Renaming project file...");
                
                // Find all .ipj files in the project directory
                var projectFiles = Directory.GetFiles(projectPath, "*.ipj", SearchOption.TopDirectoryOnly);
                
                foreach (var projectFile in projectFiles)
                {
                    try
                    {
                        var fileName = System.IO.Path.GetFileNameWithoutExtension(projectFile);
                        var fileExt = System.IO.Path.GetExtension(projectFile);
                        var directory = System.IO.Path.GetDirectoryName(projectFile);
                        
                        // Check if the file name starts with the old prefix
                        if (fileName.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            // Generate new file name with new prefix
                            var newFileName = fileName.Replace(oldPrefix, newPrefix, StringComparison.OrdinalIgnoreCase);
                            
                            if (directory != null)
                            {
                                var newFilePath = System.IO.Path.Combine(directory, newFileName + fileExt);
                                
                                // Check if the new file already exists
                                if (System.IO.File.Exists(newFilePath))
                                {
                                    Console.WriteLine($"  Skipping {System.IO.Path.GetFileName(projectFile)} - target file already exists");
                                    continue;
                                }
                                
                                // Rename the project file with retry logic for file locks
                                bool renameSuccess = false;
                                for (int retry = 0; retry < 3; retry++)
                                {
                                    try
                                    {
                                        System.IO.File.Move(projectFile, newFilePath);
                                        renameSuccess = true;
                                        break;
                                    }
                                    catch (IOException ioEx) when (ioEx.Message.Contains("being used by another process"))
                                    {
                                        Console.WriteLine($"  Project file locked, retrying in 1 second... (attempt {retry + 1}/3)");
                                        Thread.Sleep(1000);
                                    }
                                }
                                
                                if (renameSuccess)
                                {
                                    Console.WriteLine($"  Renamed project file: {System.IO.Path.GetFileName(projectFile)} -> {System.IO.Path.GetFileName(newFilePath)}");
                                    
                                    // Add to files to delete list (original file)
                                    result.FilesToDelete.Add(projectFile);
                                    
                                    result.UpdatedProjectFiles.Add($"Renamed: {System.IO.Path.GetFileName(projectFile)} -> {System.IO.Path.GetFileName(newFilePath)}");
                                }
                                else
                                {
                                    Console.WriteLine($"  Failed to rename project file {System.IO.Path.GetFileName(projectFile)} - file still locked");
                                    result.FailedProjectFiles.Add($"Failed to rename {System.IO.Path.GetFileName(projectFile)} - file locked");
                                }
                            }
                            else
                            {
                                Console.WriteLine($"  Could not get directory for project file: {projectFile}");
                            }
                        }
                        else
                        {
                            Console.WriteLine($"  Skipping {System.IO.Path.GetFileName(projectFile)} - does not start with old prefix");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  Error renaming project file {System.IO.Path.GetFileName(projectFile)}: {ex.Message}");
                        result.FailedProjectFiles.Add($"Failed to rename {System.IO.Path.GetFileName(projectFile)}: {ex.Message}");
                    }
                }
                
                Console.WriteLine("Project file renaming completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error during project file renaming: {ex.Message}");
                result.ErrorMessage = $"Project file renaming error: {ex.Message}";
            }
        }

        /// <summary>
        /// Discovers all drawing files in the specified directory.
        /// </summary>
        private List<string> DiscoverDrawingFiles(string drawingsPath)
        {
            try
            {
                return Directory.GetFiles(drawingsPath, "*.idw", SearchOption.TopDirectoryOnly)
                    .Concat(Directory.GetFiles(drawingsPath, "*.dwg", SearchOption.TopDirectoryOnly))
                    .Where(file => !string.IsNullOrEmpty(file))
                    .OrderBy(file => file)
                    .ToList();
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error discovering drawing files: {ex.Message}");
                return new List<string>();
            }
        }

        /// <summary>
        /// Builds a mapping of old file paths to new file paths for drawing reference updates.
        /// </summary>
        private Dictionary<string, string> BuildFileMappingForDrawingUpdates(string modelPath, string oldPrefix, string newPrefix)
        {
            var fileMapping = new Dictionary<string, string>();

            try
            {
                // Get all files in the model directory and subdirectories
                var allFiles = Directory.GetFiles(modelPath, "*.*", SearchOption.AllDirectories)
                    .Where(file => IsInventorFile(file))
                    .ToList();

                foreach (var file in allFiles)
                {
                    var fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                    var fileExt = System.IO.Path.GetExtension(file);
                    var directory = System.IO.Path.GetDirectoryName(file);

                    // Check if the file name starts with the old prefix
                    if (fileName.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        // Generate new file name with new prefix
                        var newFileName = fileName.Replace(oldPrefix, newPrefix, StringComparison.OrdinalIgnoreCase);

                        if (directory != null)
                        {
                            var newFilePath = System.IO.Path.Combine(directory, newFileName + fileExt);

                            // Add multiple path formats to handle different reference types
                            fileMapping[file] = newFilePath;

                            // Add just the filename mapping for cases where drawings reference by filename only
                            var oldFileName = System.IO.Path.GetFileName(file);
                            var newFileNameWithExt = newFileName + fileExt;
                            fileMapping[oldFileName] = newFilePath; // Use full path instead of just filename

                            // Add relative path mapping
                            var relativePath = System.IO.Path.GetRelativePath(modelPath, file);
                            var newRelativePath = System.IO.Path.GetRelativePath(modelPath, newFilePath);
                            fileMapping[relativePath] = newRelativePath;

                            // Add normalized path mappings (handle different path separators)
                            var normalizedOldPath = file.Replace('\\', '/');
                            var normalizedNewPath = newFilePath.Replace('\\', '/');
                            fileMapping[normalizedOldPath] = normalizedNewPath;

                            Console.WriteLine($"Mapped: {System.IO.Path.GetFileName(file)} -> {System.IO.Path.GetFileName(newFilePath)}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error building file mapping: {ex.Message}");
            }

            return fileMapping;
        }

        /// <summary>
        /// Checks if a file is an Inventor file based on its extension.
        /// </summary>
        private bool IsInventorFile(string filePath)
        {
            var extension = System.IO.Path.GetExtension(filePath).ToLowerInvariant();
            return extension == ".iam" || extension == ".ipt" || extension == ".ipn" || extension == ".idw" || extension == ".dwg";
        }

        /// <summary>
        /// Updates drawing references to point to renamed files.
        /// </summary>
        private void UpdateDrawingReferences(List<string> drawingFiles, Dictionary<string, string> fileMapping, DrawingUpdateResult result, string oldPrefix, string newPrefix)
        {
            var inventorApp = GetInventorApplication();
            int updatedDrawings = 0;
            int failedDrawings = 0;
            var processedReferences = new HashSet<string>(); // Track processed references to avoid duplicates

            foreach (var drawingFile in drawingFiles)
            {
                var drawingFileName = System.IO.Path.GetFileName(drawingFile);
                Console.WriteLine($"Processing drawing: {drawingFileName}");

                DrawingDocument? drawingDoc = null;
                try
                {
                    drawingDoc = (DrawingDocument)inventorApp.Documents.Open(drawingFile, false);
                    bool drawingUpdated = false;
                    int referencesUpdated = 0;
                    int referencesFailed = 0;

                    // Process all sheets in the drawing
                    var sheets = drawingDoc.Sheets;
                    foreach (Sheet sheet in sheets)
                    {
                        var views = sheet.DrawingViews;
                        foreach (DrawingView view in views)
                        {
                            try
                            {
                                var referencedPath = view.ReferencedDocumentDescriptor?.FullDocumentName;
                                if (string.IsNullOrEmpty(referencedPath))
                                    continue;

                                // Create a unique key for this reference to avoid processing duplicates
                                var referenceKey = $"{drawingFileName}:{referencedPath}";
                                if (processedReferences.Contains(referenceKey))
                                {
                                    Console.WriteLine($"  Skipping duplicate reference: {System.IO.Path.GetFileName(referencedPath)}");
                                    continue;
                                }
                                processedReferences.Add(referenceKey);

                                // Check if this reference needs to be updated
                                Console.WriteLine($"  Checking reference: {referencedPath}");

                                // Try to find a mapping for this reference
                                string? newPath = null;
                                string? mappingKey = null;

                                // Try exact path match first
                                if (fileMapping.ContainsKey(referencedPath))
                                {
                                    newPath = fileMapping[referencedPath];
                                    mappingKey = referencedPath;
                                }
                                // Try normalized path match (handle different path separators)
                                else
                                {
                                    var normalizedPath = referencedPath.Replace('\\', '/');
                                    if (fileMapping.ContainsKey(normalizedPath))
                                    {
                                        newPath = fileMapping[normalizedPath];
                                        mappingKey = normalizedPath;
                                    }
                                }

                                // Try filename only match - but we need to find the full path
                                if (newPath == null)
                                {
                                    var fileName = System.IO.Path.GetFileName(referencedPath);
                                    if (fileMapping.ContainsKey(fileName))
                                    {
                                        newPath = fileMapping[fileName]; // This now contains the full path
                                        mappingKey = fileName;
                                    }
                                }

                                // Try case-insensitive filename match
                                if (newPath == null)
                                {
                                    var fileName = System.IO.Path.GetFileName(referencedPath);
                                    var caseInsensitiveMapping = fileMapping.FirstOrDefault(kvp =>
                                        string.Equals(System.IO.Path.GetFileName(kvp.Key), fileName, StringComparison.OrdinalIgnoreCase));

                                    if (caseInsensitiveMapping.Key != null)
                                    {
                                        newPath = caseInsensitiveMapping.Value;
                                        mappingKey = caseInsensitiveMapping.Key;
                                    }
                                }

                                if (newPath != null)
                                {
                                    Console.WriteLine($"  Found mapping: {mappingKey} -> {System.IO.Path.GetFileName(newPath)}");
                                    Console.WriteLine($"  Full new path: {newPath}");
                                    Console.WriteLine($"  Updating reference: {System.IO.Path.GetFileName(referencedPath)} -> {System.IO.Path.GetFileName(newPath)}");

                                    // Try to update the reference
                                    if (UpdateDrawingReference(drawingDoc, view, referencedPath, newPath, drawingFileName))
                                    {
                                        drawingUpdated = true;
                                        referencesUpdated++;
                                        result.UpdatedReferences.Add($"{drawingFileName}: {System.IO.Path.GetFileName(referencedPath)} -> {System.IO.Path.GetFileName(newPath)}");
                                        Console.WriteLine($"  ✓ Reference updated successfully");
                                    }
                                    else
                                    {
                                        referencesFailed++;
                                        result.FailedReferences.Add($"{drawingFileName}: Failed to update {System.IO.Path.GetFileName(referencedPath)}");
                                        Console.WriteLine($"  ✗ Reference update failed");
                                    }
                                }
                                else
                                {
                                    Console.WriteLine($"  Reference does not need updating: {System.IO.Path.GetFileName(referencedPath)}");
                                    Console.WriteLine($"  (No mapping found for this reference)");
                                }
                            }
                            catch (Exception viewEx)
                            {
                                Console.WriteLine($"  Error processing view: {viewEx.Message}");
                                result.FailedReferences.Add($"{drawingFileName}: View processing error - {viewEx.Message}");
                                referencesFailed++;
                            }
                        }
                    }

                    if (drawingUpdated)
                    {
                        Console.WriteLine($"  Saving drawing with {referencesUpdated} updated references...");
                        drawingDoc.Save();
                        updatedDrawings++;
                        Console.WriteLine($"  ✓ Drawing updated successfully: {drawingFileName} ({referencesUpdated} references updated, {referencesFailed} failed)");
                    }
                    else
                    {
                        Console.WriteLine($"  - No updates needed for: {drawingFileName}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  ✗ Error processing drawing {drawingFileName}: {ex.Message}");
                    result.FailedDrawings.Add($"{drawingFileName}: {ex.Message}");
                    failedDrawings++;
                }
                finally
                {
                    if (drawingDoc != null)
                    {
                        try
                        {
                            drawingDoc.Close(false);
                            Marshal.ReleaseComObject(drawingDoc);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Error closing drawing document: {ex.Message}");
                        }
                    }
                }
            }

            result.UpdatedDrawings = updatedDrawings;
            result.FailedDrawingsCount = failedDrawings;
            Console.WriteLine($"Drawing update summary: {updatedDrawings} updated, {failedDrawings} failed");
        }

        /// <summary>
        /// Updates a specific drawing reference using document descriptor approach.
        /// </summary>
        private bool UpdateDrawingReference(DrawingDocument drawingDoc, DrawingView view, string oldPath, string newPath, string drawingFileName)
        {
            try
            {
                // Check if the new file exists
                if (!System.IO.File.Exists(newPath))
                {
                    Console.WriteLine($"    New file does not exist: {System.IO.Path.GetFileName(newPath)}");
                    return false;
                }

                var refDocDesc = view.ReferencedDocumentDescriptor;
                if (refDocDesc == null)
                {
                    Console.WriteLine($"    No document descriptor found for view");
                    return false;
                }

                // Check current reference path
                var currentPath = refDocDesc.FullDocumentName;
                Console.WriteLine($"    Current reference: {System.IO.Path.GetFileName(currentPath)}");
                Console.WriteLine($"    Target reference: {System.IO.Path.GetFileName(newPath)}");

                // If already pointing to the correct file, no update needed
                if (currentPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine($"    Reference already correct: {System.IO.Path.GetFileName(newPath)}");
                    return true;
                }

                // Method 1: Use Document Descriptor ReplaceReference approach
                try
                {
                    Console.WriteLine($"    Attempting Document Descriptor ReplaceReference method...");

                    var inventorApp = GetInventorApplication();

                    // Close the old document if it's open
                    var oldDoc = refDocDesc.ReferencedDocument;
                    if (oldDoc != null)
                    {
                        Console.WriteLine($"    Closing old document: {System.IO.Path.GetFileName(currentPath)}");
                        ((Inventor.Document)oldDoc).Close(false);
                        Marshal.ReleaseComObject(oldDoc);
                        Thread.Sleep(1000);
                    }

                    // Try to use the document descriptor's ReplaceReference method
                    // This is the proper Inventor API way to replace references
                    try
                    {
                        // Use the document descriptor to replace the reference
                        // The ReplaceReference method should be available on the document descriptor
                        refDocDesc.ReferencedFileDescriptor.ReplaceReference(newPath);
                        Console.WriteLine($"    ReplaceReference called successfully");

                        // Force the drawing to update
                        drawingDoc.Update();
                        Thread.Sleep(2000);

                        // Check if the reference was updated
                        var updatedPath = view.ReferencedDocumentDescriptor?.FullDocumentName;
                        Console.WriteLine($"    Reference after ReplaceReference: {System.IO.Path.GetFileName(updatedPath)}");

                        if (updatedPath != null && updatedPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine($"    ✓ Reference successfully updated using Document Descriptor ReplaceReference: {System.IO.Path.GetFileName(newPath)}");
                            return true;
                        }
                    }
                    catch (Exception replaceEx)
                    {
                        Console.WriteLine($"    Document Descriptor ReplaceReference failed: {replaceEx.Message}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    Document Descriptor approach failed: {ex.Message}");
                }

                // Method 2: Try using Inventor's reference management with forced path update
                try
                {
                    Console.WriteLine($"    Attempting forced path update method...");

                    // Close the old document if it's open
                    var oldDoc = refDocDesc.ReferencedDocument;
                    if (oldDoc != null)
                    {
                        Console.WriteLine($"    Closing old document: {System.IO.Path.GetFileName(currentPath)}");
                        ((Inventor.Document)oldDoc).Close(false);
                        Marshal.ReleaseComObject(oldDoc);
                        Thread.Sleep(1000);
                    }

                    // Force the drawing to update and look for the new file
                    drawingDoc.Update();
                    Thread.Sleep(2000);
                    drawingDoc.Save();
                    Thread.Sleep(1000);

                    // Check if the reference was updated
                    var updatedPath = view.ReferencedDocumentDescriptor?.FullDocumentName;
                    Console.WriteLine($"    Reference after forced update: {System.IO.Path.GetFileName(updatedPath)}");

                    if (updatedPath != null && updatedPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"    ✓ Reference successfully updated using forced update: {System.IO.Path.GetFileName(newPath)}");
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    Forced update approach failed: {ex.Message}");
                }

                // Method 3: Try using Inventor's document-level reference update
                try
                {
                    Console.WriteLine($"    Attempting document-level reference update...");

                    var inventorApp = GetInventorApplication();
                    inventorApp.SilentOperation = true;

                    // Force update the entire drawing document
                    drawingDoc.Update();
                    Thread.Sleep(2000);
                    drawingDoc.Save();
                    Thread.Sleep(1000);

                    // Check if the reference was updated
                    var updatedPath = view.ReferencedDocumentDescriptor?.FullDocumentName;
                    Console.WriteLine($"    Reference after document-level update: {System.IO.Path.GetFileName(updatedPath)}");

                    if (updatedPath != null && updatedPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"    ✓ Reference successfully updated using document-level update: {System.IO.Path.GetFileName(newPath)}");
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"    Document-level update approach failed: {ex.Message}");
                }

                Console.WriteLine($"    ✗ Reference update failed for: {System.IO.Path.GetFileName(oldPath)} -> {System.IO.Path.GetFileName(newPath)}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"    Error updating drawing reference: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Updates project files to reflect renamed files.
        /// </summary>
        private void UpdateProjectFiles(string projectPath, string oldPrefix, string newPrefix, DrawingUpdateResult result)
        {
            try
            {
                var projectFiles = Directory.GetFiles(projectPath, "*.ipj", SearchOption.TopDirectoryOnly);

                foreach (var projectFile in projectFiles)
                {
                    var projectFileName = System.IO.Path.GetFileName(projectFile);
                    Console.WriteLine($"Processing project file: {projectFileName}");

                    try
                    {
                        var projectContent = System.IO.File.ReadAllText(projectFile);
                        var originalContent = projectContent;
                        bool contentUpdated = false;

                        // Update file references in the project file
                        var allFiles = Directory.GetFiles(projectPath, "*.*", SearchOption.AllDirectories)
                            .Where(file => IsInventorFile(file))
                            .ToList();

                        foreach (var file in allFiles)
                        {
                            var fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                            if (fileName.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                            {
                                var newFileName = fileName.Replace(oldPrefix, newPrefix, StringComparison.OrdinalIgnoreCase);
                                var fileExt = System.IO.Path.GetExtension(file);
                                var directory = System.IO.Path.GetDirectoryName(file);
                                if (directory != null)
                                {
                                    var newFilePath = System.IO.Path.Combine(directory, newFileName + fileExt);

                                    if (System.IO.File.Exists(newFilePath))
                                    {
                                        // Update absolute paths
                                        if (projectContent.Contains(file, StringComparison.OrdinalIgnoreCase))
                                        {
                                            projectContent = projectContent.Replace(file, newFilePath, StringComparison.OrdinalIgnoreCase);
                                            contentUpdated = true;
                                            Console.WriteLine($"  Updated absolute path: {System.IO.Path.GetFileName(file)} -> {System.IO.Path.GetFileName(newFilePath)}");
                                        }

                                        // Update relative paths
                                        var relativePath = System.IO.Path.GetRelativePath(projectPath, file);
                                        var newRelativePath = System.IO.Path.GetRelativePath(projectPath, newFilePath);
                                        if (projectContent.Contains(relativePath, StringComparison.OrdinalIgnoreCase))
                                        {
                                            projectContent = projectContent.Replace(relativePath, newRelativePath, StringComparison.OrdinalIgnoreCase);
                                            contentUpdated = true;
                                            Console.WriteLine($"  Updated relative path: {relativePath} -> {newRelativePath}");
                                        }
                                    }
                                }
                            }
                        }

                        if (contentUpdated)
                        {
                            System.IO.File.WriteAllText(projectFile, projectContent);
                            result.UpdatedProjectFiles.Add(projectFileName);
                            Console.WriteLine($"  ✓ Project file updated: {projectFileName}");
                        }
                        else
                        {
                            Console.WriteLine($"  - No updates needed for project file: {projectFileName}");
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"  ✗ Error processing project file {projectFileName}: {ex.Message}");
                        result.FailedProjectFiles.Add($"{projectFileName}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error updating project files: {ex.Message}");
                result.ErrorMessage = ex.Message;
            }
        }

        /// <summary>
        /// Result class for drawing update operations.
        /// </summary>
        public class DrawingUpdateResult
        {
            public List<string> FilesToDelete { get; set; } = new();
            public List<string> UpdatedReferences { get; set; } = new();
            public List<string> FailedReferences { get; set; } = new();
            public List<string> FailedDrawings { get; set; } = new();
            public List<string> UpdatedProjectFiles { get; set; } = new();
            public List<string> FailedProjectFiles { get; set; } = new();
            public int UpdatedDrawings { get; set; }
            public int FailedDrawingsCount { get; set; }
            public string ErrorMessage { get; set; } = "";
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

        /// <summary>
        /// Forces reference updates in drawing files by using a more aggressive approach
        /// </summary>
        private bool ForceReferenceUpdate(DrawingDocument drawingDoc, DrawingView view, string oldPath, string newPath, string drawingFileName)
        {
            try
            {
                Console.WriteLine($"Attempting to force reference update: {System.IO.Path.GetFileName(oldPath)} -> {System.IO.Path.GetFileName(newPath)}");
                Console.WriteLine($"  Old path: {oldPath}");
                Console.WriteLine($"  New path: {newPath}");
                Console.WriteLine($"  Old file exists: {System.IO.File.Exists(oldPath)}");
                Console.WriteLine($"  New file exists: {System.IO.File.Exists(newPath)}");

                // Method 1: Try to directly update the reference path using Inventor's ReplaceReference
                try
                {
                    // Check if the new file exists
                    if (!System.IO.File.Exists(newPath))
                    {
                        Console.WriteLine($"Target file does not exist: {newPath}");
                        return false;
                    }

                    // Get the referenced document descriptor
                    var refDocDesc = view.ReferencedDocumentDescriptor;
                    if (refDocDesc != null)
                    {
                        // Try to use Inventor's ReplaceReference method
                        try
                        {
                            // Close the old reference first
                            var oldDoc = refDocDesc.ReferencedDocument;
                            if (oldDoc != null)
                            {
                                ((Inventor.Document)oldDoc).Close(false);
                                Thread.Sleep(500);
                            }

                            // Force the drawing to update
                            drawingDoc.Update();
                            Thread.Sleep(1000);
                            drawingDoc.Save();
                            Thread.Sleep(1000);

                            // Check if reference was updated
                            var currentPath = view.ReferencedDocumentDescriptor?.FullDocumentName;
                            Console.WriteLine($"  Current reference path after close/update: {currentPath}");
                            Console.WriteLine($"  Expected new path: {newPath}");

                            if (currentPath != null && currentPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine($"Reference successfully updated using close/update method");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Close/update method failed: {ex.Message}");
                        }

                        // If close/update didn't work, try to force the reference to the new path
                        try
                        {
                            // Force the drawing to look for the new file
                            drawingDoc.Update();
                            Thread.Sleep(1000);

                            // Check if reference was updated
                            var currentPath = view.ReferencedDocumentDescriptor?.FullDocumentName;
                            Console.WriteLine($"  Current reference path after force update: {currentPath}");
                            Console.WriteLine($"  Expected new path: {newPath}");

                            if (currentPath != null && currentPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine($"Reference successfully updated using force update method");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Force update method failed: {ex.Message}");
                        }
                    }

                    Console.WriteLine($"Direct reference update methods did not work");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Direct reference update failed: {ex.Message}");
                }

                // Method 2: Try using Inventor's reference management
                try
                {
                    // Get the referenced document descriptor
                    var refDocDesc = view.ReferencedDocumentDescriptor;
                    if (refDocDesc != null)
                    {
                        // Try to update the reference using Inventor's built-in methods
                        drawingDoc.Update();
                        Thread.Sleep(1000);

                        // Force a save to ensure changes are persisted
                        drawingDoc.Save();
                        Thread.Sleep(1000);

                        // Check if reference was updated
                        var updatedPath = view.ReferencedDocumentDescriptor?.FullDocumentName;
                        if (updatedPath != null && updatedPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                        {
                            Console.WriteLine($"Reference successfully updated using Inventor reference management");
                            return true;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Inventor reference management failed: {ex.Message}");
                }

                // Method 3: Try using Inventor's ReplaceReference method on the view
                try
                {
                    // Get the referenced document descriptor
                    var refDocDesc = view.ReferencedDocumentDescriptor;
                    if (refDocDesc != null)
                    {
                        // Try to replace the reference directly
                        try
                        {
                            // Use Inventor's ReplaceReference method if available
                            if (refDocDesc.ReferencedDocument != null)
                            {
                                // Close the old document
                                ((Inventor.Document)refDocDesc.ReferencedDocument).Close(false);
                                Thread.Sleep(500);
                            }

                            // Force the drawing to update and look for the new file
                            drawingDoc.Update();
                            Thread.Sleep(1000);
                            drawingDoc.Save();
                            Thread.Sleep(1000);

                            // Check if reference was updated
                            var updatedPath = view.ReferencedDocumentDescriptor?.FullDocumentName;
                            Console.WriteLine($"  Current reference path after ReplaceReference: {updatedPath}");
                            Console.WriteLine($"  Expected new path: {newPath}");

                            if (updatedPath != null && updatedPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                            {
                                Console.WriteLine($"Reference successfully updated using ReplaceReference method");
                                return true;
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"ReplaceReference method failed: {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ReplaceReference update failed: {ex.Message}");
                }

                // Method 4: Document-level reference update
                try
                {
                    // Try to update all references in the drawing document
                    var inventorApp = GetInventorApplication();
                    inventorApp.SilentOperation = true;

                    // Force update the entire drawing
                    drawingDoc.Update();
                    Thread.Sleep(1000);
                    drawingDoc.Save();
                    Thread.Sleep(1000);

                    // Check if reference was updated
                    var updatedPath = view.ReferencedDocumentDescriptor?.FullDocumentName;
                    if (updatedPath != null && updatedPath.Equals(newPath, StringComparison.OrdinalIgnoreCase))
                    {
                        Console.WriteLine($"Reference successfully updated using document-level update");
                        return true;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Document-level update failed: {ex.Message}");
                }

                Console.WriteLine($"All reference update methods failed for {System.IO.Path.GetFileName(oldPath)} -> {System.IO.Path.GetFileName(newPath)}");
                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Force reference update failed: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Validates that drawing links resolve correctly after updates
        /// </summary>
        private void ValidateDrawingLinks(List<string> drawingFiles, Dictionary<string, string> fileMapping, DrawingUpdateResult result, string oldPrefix = "ABC099001", string newPrefix = "Test1232")
        {
            Console.WriteLine($"=== Validating Drawing Links ===");

            var inventorApp = GetInventorApplication();
            inventorApp.SilentOperation = true;
            inventorApp.Visible = false;

            int validatedDrawings = 0;
            int failedValidations = 0;

            foreach (var drawingFile in drawingFiles)
            {
                string drawingFileName = System.IO.Path.GetFileName(drawingFile);
                Console.WriteLine($"Validating links in: {drawingFileName}");

                DrawingDocument? drawingDoc = null;
                try
                {
                    drawingDoc = (DrawingDocument)inventorApp.Documents.Open(drawingFile, false);

                    bool allLinksValid = true;
                    var sheets = drawingDoc.Sheets;

                    foreach (Sheet sheet in sheets)
                    {
                        var views = sheet.DrawingViews;

                        foreach (DrawingView view in views)
                        {
                            try
                            {
                                string? referencedPath = view.ReferencedDocumentDescriptor?.FullDocumentName;

                                if (string.IsNullOrEmpty(referencedPath))
                                {
                                    Console.WriteLine($"  Warning: No referenced document found for view in {drawingFileName}");
                                    continue;
                                }

                                // Check if the referenced file exists
                                if (!System.IO.File.Exists(referencedPath))
                                {
                                    Console.WriteLine($"  ✗ Error: Referenced file not found: {System.IO.Path.GetFileName(referencedPath)}");
                                    allLinksValid = false;
                                    result.FailedDrawings.Add($"Missing reference: {drawingFileName} -> {System.IO.Path.GetFileName(referencedPath)}");
                                }
                                else
                                {
                                    // Check if this reference should have been updated
                                    if (fileMapping.ContainsKey(referencedPath))
                                    {
                                        string expectedNewPath = fileMapping[referencedPath];
                                        if (!referencedPath.Equals(expectedNewPath, StringComparison.OrdinalIgnoreCase))
                                        {
                                            Console.WriteLine($"  ✗ INCORRECT REFERENCE: {System.IO.Path.GetFileName(referencedPath)} (should be {System.IO.Path.GetFileName(expectedNewPath)})");
                                            allLinksValid = false;
                                            result.UpdatedReferences.Add($"INCORRECT_REFERENCE: {drawingFileName} - {System.IO.Path.GetFileName(referencedPath)} should be {System.IO.Path.GetFileName(expectedNewPath)}");
                                        }
                                        else
                                        {
                                            Console.WriteLine($"  ✓ Reference correctly updated: {System.IO.Path.GetFileName(referencedPath)}");
                                        }
                                    }
                                    else
                                    {
                                        // Check if this is an old reference that should have been updated
                                        string fileName = System.IO.Path.GetFileName(referencedPath);
                                        if (fileName.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                                        {
                                            // This is an old reference that should have been updated
                                            string expectedNewFileName = fileName.Replace(oldPrefix, newPrefix, StringComparison.OrdinalIgnoreCase);
                                            string expectedNewPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(referencedPath) ?? "", expectedNewFileName);

                                            if (System.IO.File.Exists(expectedNewPath))
                                            {
                                                Console.WriteLine($"  ✗ OLD REFERENCE DETECTED: {System.IO.Path.GetFileName(referencedPath)} (should be {System.IO.Path.GetFileName(expectedNewPath)})");
                                                allLinksValid = false;
                                                result.UpdatedReferences.Add($"OLD_REFERENCE_DETECTED: {drawingFileName} - {System.IO.Path.GetFileName(referencedPath)} should be {System.IO.Path.GetFileName(expectedNewPath)}");
                                            }
                                            else
                                            {
                                                Console.WriteLine($"  ✓ Reference valid: {System.IO.Path.GetFileName(referencedPath)}");
                                            }
                                        }
                                        else
                                        {
                                            Console.WriteLine($"  ✓ Reference valid: {System.IO.Path.GetFileName(referencedPath)}");
                                        }
                                    }
                                }
                            }
                            catch (Exception viewEx)
                            {
                                Console.WriteLine($"  Error validating view in {drawingFileName}: {viewEx.Message}");
                                allLinksValid = false;
                            }
                        }
                    }

                    if (allLinksValid)
                    {
                        Console.WriteLine($"  ✓ All links valid in {drawingFileName}");
                        validatedDrawings++;
                    }
                    else
                    {
                        Console.WriteLine($"  ✗ Some links invalid in {drawingFileName}");
                        failedValidations++;
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error validating drawing {drawingFileName}: {ex.Message}");
                    result.FailedDrawings.Add($"Link validation error: {drawingFileName} - {ex.Message}");
                    failedValidations++;
                }
                finally
                {
                    if (drawingDoc != null)
                    {
                        try
                        {
                            drawingDoc.Close(false);
                            Marshal.ReleaseComObject(drawingDoc);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Error closing drawing document: {ex.Message}");
                        }
                    }
                }
            }

            Console.WriteLine($"Link validation summary: {validatedDrawings} drawings with valid links, {failedValidations} drawings with issues");
        }

        /// <summary>
        /// Validates drawing files and their references for integrity
        /// </summary>
        private void ValidateDrawingFiles(List<string> drawingFiles, Dictionary<string, string> fileMapping, DrawingUpdateResult result)
        {
            Console.WriteLine($"=== Validating Drawing Files ===");

            foreach (var drawingFile in drawingFiles)
            {
                try
                {
                    string drawingFileName = System.IO.Path.GetFileName(drawingFile);
                    string extension = System.IO.Path.GetExtension(drawingFile).ToLower();

                    // Validate file extension
                    if (extension != ".idw" && extension != ".dwg")
                    {
                        Console.WriteLine($"Warning: Invalid drawing file extension: {drawingFileName}");
                        result.FailedDrawings.Add($"Invalid extension: {drawingFileName}");
                        continue;
                    }

                    // Check if file is accessible
                    if (IsFileLocked(drawingFile))
                    {
                        Console.WriteLine($"Warning: Drawing file is locked: {drawingFileName}");
                        result.FailedDrawings.Add($"File locked: {drawingFileName}");
                        continue;
                    }

                    // Check file size (basic integrity check)
                    var fileInfo = new System.IO.FileInfo(drawingFile);
                    if (fileInfo.Length == 0)
                    {
                        Console.WriteLine($"Warning: Drawing file is empty: {drawingFileName}");
                        result.FailedDrawings.Add($"Empty file: {drawingFileName}");
                        continue;
                    }

                    // Check if file is too large (potential corruption indicator)
                    if (fileInfo.Length > 500 * 1024 * 1024) // 500MB
                    {
                        Console.WriteLine($"Warning: Drawing file is unusually large: {drawingFileName} ({fileInfo.Length / (1024 * 1024)}MB)");
                    }

                    Console.WriteLine($"Drawing file validation passed: {drawingFileName}");
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error validating drawing file {System.IO.Path.GetFileName(drawingFile)}: {ex.Message}");
                    result.FailedDrawings.Add($"Validation error: {System.IO.Path.GetFileName(drawingFile)} - {ex.Message}");
                }
            }

            Console.WriteLine($"=== Drawing File Validation Completed ===");
        }

        /// <summary>
        /// Updates the content of .ipj project files to reflect new file paths and references
        /// </summary>
        private List<string> UpdateProjectFileContent(List<string> projectFiles, Dictionary<string, string> fileMapping, string oldPrefix, string newPrefix, DrawingUpdateResult result)
        {
            var updatedProjectFiles = new List<string>();

            foreach (var projectFile in projectFiles)
            {
                try
                {
                    Console.WriteLine($"Updating project file content: {System.IO.Path.GetFileName(projectFile)}");

                    // Read the project file content
                    string projectContent = System.IO.File.ReadAllText(projectFile);
                    string originalContent = projectContent;

                    // Update file references in the project file
                    bool contentUpdated = false;

                    // Update references to model files
                    foreach (var mapping in fileMapping)
                    {
                        string oldPath = mapping.Key;
                        string newPath = mapping.Value;

                        // Convert paths to relative paths if they're in the same directory structure
                        string oldRelativePath = GetRelativePath(projectFile, oldPath);
                        string newRelativePath = GetRelativePath(projectFile, newPath);

                        // Update both absolute and relative paths
                        if (projectContent.Contains(oldPath, StringComparison.OrdinalIgnoreCase))
                        {
                            projectContent = projectContent.Replace(oldPath, newPath, StringComparison.OrdinalIgnoreCase);
                            contentUpdated = true;
                            Console.WriteLine($"  Updated absolute path: {System.IO.Path.GetFileName(oldPath)} -> {System.IO.Path.GetFileName(newPath)}");
                        }

                        if (projectContent.Contains(oldRelativePath, StringComparison.OrdinalIgnoreCase))
                        {
                            projectContent = projectContent.Replace(oldRelativePath, newRelativePath, StringComparison.OrdinalIgnoreCase);
                            contentUpdated = true;
                            Console.WriteLine($"  Updated relative path: {System.IO.Path.GetFileName(oldPath)} -> {System.IO.Path.GetFileName(newPath)}");
                        }
                    }

                    // Update drawing file references
                    var drawingFiles = System.IO.Directory.GetFiles(System.IO.Path.GetDirectoryName(projectFile) ?? "", "*.idw", System.IO.SearchOption.TopDirectoryOnly)
                        .Concat(System.IO.Directory.GetFiles(System.IO.Path.GetDirectoryName(projectFile) ?? "", "*.dwg", System.IO.SearchOption.TopDirectoryOnly))
                        .ToList();

                    foreach (var drawingFile in drawingFiles)
                    {
                        string drawingFileName = System.IO.Path.GetFileNameWithoutExtension(drawingFile);
                        string drawingExtension = System.IO.Path.GetExtension(drawingFile);

                        if (drawingFileName.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                        {
                            string newDrawingFileName = drawingFileName.Replace(oldPrefix, newPrefix, StringComparison.OrdinalIgnoreCase);
                            string oldDrawingPath = drawingFile;
                            string newDrawingPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(drawingFile) ?? "", newDrawingFileName + drawingExtension);

                            if (System.IO.File.Exists(newDrawingPath))
                            {
                                string oldRelativeDrawingPath = GetRelativePath(projectFile, oldDrawingPath);
                                string newRelativeDrawingPath = GetRelativePath(projectFile, newDrawingPath);

                                if (projectContent.Contains(oldDrawingPath, StringComparison.OrdinalIgnoreCase))
                                {
                                    projectContent = projectContent.Replace(oldDrawingPath, newDrawingPath, StringComparison.OrdinalIgnoreCase);
                                    contentUpdated = true;
                                    Console.WriteLine($"  Updated drawing absolute path: {System.IO.Path.GetFileName(oldDrawingPath)} -> {System.IO.Path.GetFileName(newDrawingPath)}");
                                }

                                if (projectContent.Contains(oldRelativeDrawingPath, StringComparison.OrdinalIgnoreCase))
                                {
                                    projectContent = projectContent.Replace(oldRelativeDrawingPath, newRelativeDrawingPath, StringComparison.OrdinalIgnoreCase);
                                    contentUpdated = true;
                                    Console.WriteLine($"  Updated drawing relative path: {System.IO.Path.GetFileName(oldDrawingPath)} -> {System.IO.Path.GetFileName(newDrawingPath)}");
                                }
                            }
                        }
                    }

                    // Update project file name references
                    string projectFileName = System.IO.Path.GetFileNameWithoutExtension(projectFile);
                    if (projectFileName.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        string newProjectFileName = projectFileName.Replace(oldPrefix, newPrefix, StringComparison.OrdinalIgnoreCase);
                        string newProjectPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(projectFile) ?? "", newProjectFileName + System.IO.Path.GetExtension(projectFile));

                        if (System.IO.File.Exists(newProjectPath))
                        {
                            if (projectContent.Contains(projectFile, StringComparison.OrdinalIgnoreCase))
                            {
                                projectContent = projectContent.Replace(projectFile, newProjectPath, StringComparison.OrdinalIgnoreCase);
                                contentUpdated = true;
                                Console.WriteLine($"  Updated project file reference: {System.IO.Path.GetFileName(projectFile)} -> {System.IO.Path.GetFileName(newProjectPath)}");
                            }
                        }
                    }

                    // Save the updated project file if changes were made
                    if (contentUpdated)
                    {
                        // Create a backup of the original file
                        string backupFile = projectFile + ".backup";
                        System.IO.File.Copy(projectFile, backupFile, true);

                        // Write the updated content
                        System.IO.File.WriteAllText(projectFile, projectContent);

                        updatedProjectFiles.Add(System.IO.Path.GetFileName(projectFile));
                        Console.WriteLine($"  Successfully updated project file content: {System.IO.Path.GetFileName(projectFile)}");

                        // Clean up backup file after successful update
                        try
                        {
                            System.IO.File.Delete(backupFile);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"  Warning: Could not delete backup file {backupFile}: {ex.Message}");
                        }
                    }
                    else
                    {
                        Console.WriteLine($"  No updates needed for project file: {System.IO.Path.GetFileName(projectFile)}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"  Error updating project file {System.IO.Path.GetFileName(projectFile)}: {ex.Message}");
                    result.FailedProjectFiles.Add($"Content update failed: {System.IO.Path.GetFileName(projectFile)} - {ex.Message}");
                }
            }

            return updatedProjectFiles;
        }

        /// <summary>
        /// Gets a relative path from one file to another
        /// </summary>
        private string GetRelativePath(string fromPath, string toPath)
        {
            try
            {
                var fromUri = new Uri(System.IO.Path.GetFullPath(fromPath));
                var toUri = new Uri(System.IO.Path.GetFullPath(toPath));

                if (fromUri.Scheme != toUri.Scheme)
                {
                    return toPath; // Path can't be made relative
                }

                var relativeUri = fromUri.MakeRelativeUri(toUri);
                var relativePath = Uri.UnescapeDataString(relativeUri.ToString());

                return relativePath.Replace('/', System.IO.Path.DirectorySeparatorChar);
            }
            catch
            {
                return toPath; // If we can't make it relative, return the full path
            }
        }

        /// <summary>
        /// Builds a comprehensive mapping of old file paths to new file paths for both .ipt and .iam files
        /// </summary>
        private Dictionary<string, string> BuildFileMapping(string modelPath, string oldPrefix, string newPrefix)
        {
            var fileMapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            try
            {
                // Optimized: Single directory scan for all model files
                var allFiles = System.IO.Directory.GetFiles(modelPath, "*.*", System.IO.SearchOption.TopDirectoryOnly)
                    .Where(f =>
                    {
                        var ext = System.IO.Path.GetExtension(f).ToLower();
                        return ext == ".iam" || ext == ".ipt";
                    }).ToList();

                var assemblyCount = allFiles.Count(f => System.IO.Path.GetExtension(f).ToLower() == ".iam");
                var partCount = allFiles.Count(f => System.IO.Path.GetExtension(f).ToLower() == ".ipt");
                Console.WriteLine($"Found {assemblyCount} assembly files and {partCount} part files in model directory.");

                // Debug: List all files found
                Console.WriteLine("=== DEBUG: All model files found ===");
                foreach (var file in allFiles.Take(10)) // Show first 10 files
                {
                    Console.WriteLine($"  {System.IO.Path.GetFileName(file)}");
                }
                if (allFiles.Count > 10)
                {
                    Console.WriteLine($"  ... and {allFiles.Count - 10} more files");
                }

                // Count files by prefix
                var oldPrefixFiles = allFiles.Where(f => System.IO.Path.GetFileNameWithoutExtension(f).StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase)).ToList();
                var newPrefixFiles = allFiles.Where(f => System.IO.Path.GetFileNameWithoutExtension(f).StartsWith(newPrefix, StringComparison.OrdinalIgnoreCase)).ToList();
                Console.WriteLine($"Files with old prefix '{oldPrefix}': {oldPrefixFiles.Count}");
                Console.WriteLine($"Files with new prefix '{newPrefix}': {newPrefixFiles.Count}");

                // Special case: If all files have been renamed to new prefix, create virtual mappings
                if (oldPrefixFiles.Count == 0 && newPrefixFiles.Count > 0)
                {
                    Console.WriteLine("=== All files already renamed - creating virtual mappings ===");
                    foreach (var newFile in newPrefixFiles)
                    {
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(newFile);
                        string extension = System.IO.Path.GetExtension(newFile);

                        // Create the expected old file name
                        string oldFileName = fileName.Replace(newPrefix, oldPrefix, StringComparison.OrdinalIgnoreCase);
                        string oldFilePath = System.IO.Path.Combine(modelPath, oldFileName + extension);

                        // Create virtual mapping (old path -> new path)
                        fileMapping[oldFilePath] = newFile;
                        string fileType = extension.ToLower() == ".iam" ? "assembly" : "part";
                        Console.WriteLine($"Created virtual {fileType} mapping: {oldFileName + extension} -> {System.IO.Path.GetFileName(newFile)}");
                    }

                    // Also create PROJECTS to WIP mappings
                    Console.WriteLine("=== Creating PROJECTS to WIP mappings for renamed files ===");
                    foreach (var newFile in newPrefixFiles)
                    {
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(newFile);
                        string extension = System.IO.Path.GetExtension(newFile);

                        // Create the expected old file name
                        string oldFileName = fileName.Replace(newPrefix, oldPrefix, StringComparison.OrdinalIgnoreCase);

                        // Create PROJECTS path (old path)
                        string projectsPath = modelPath.Replace("\\WIP\\", "\\PROJECTS\\");
                        string oldFilePath = System.IO.Path.Combine(projectsPath, oldFileName + extension);

                        // Create WIP path (new path)
                        string newFilePath = System.IO.Path.Combine(modelPath, fileName + extension);

                        // Only add if not already mapped
                        if (!fileMapping.ContainsKey(oldFilePath))
                        {
                            fileMapping[oldFilePath] = newFilePath;
                            string fileType = extension.ToLower() == ".iam" ? "assembly" : "part";
                            Console.WriteLine($"Created {fileType} mapping (PROJECTS->WIP): {oldFileName + extension} -> {System.IO.Path.GetFileName(newFilePath)}");
                        }
                    }
                    Console.WriteLine($"=== Virtual mappings created: {fileMapping.Count} ===");
                    return fileMapping; // Return early since we've created all mappings
                }

                // Special case: If we have both old and new files, create mappings between them
                if (oldPrefixFiles.Count > 0 && newPrefixFiles.Count > 0)
                {
                    Console.WriteLine("=== Mixed files found - creating mappings between old and new ===");
                    foreach (var oldFile in oldPrefixFiles)
                    {
                        string oldFileName = System.IO.Path.GetFileNameWithoutExtension(oldFile);
                        string extension = System.IO.Path.GetExtension(oldFile);

                        // Create the expected new file name
                        string newFileName = oldFileName.Replace(oldPrefix, newPrefix, StringComparison.OrdinalIgnoreCase);
                        string newFilePath = System.IO.Path.Combine(modelPath, newFileName + extension);

                        // Check if the corresponding new file exists
                        if (System.IO.File.Exists(newFilePath))
                        {
                            fileMapping[oldFile] = newFilePath;
                            string fileType = extension.ToLower() == ".iam" ? "assembly" : "part";
                            Console.WriteLine($"Created {fileType} mapping: {System.IO.Path.GetFileName(oldFile)} -> {System.IO.Path.GetFileName(newFilePath)}");
                        }
                    }
                    Console.WriteLine($"=== Mixed file mappings created: {fileMapping.Count} ===");

                    // Also create PROJECTS to WIP mappings for any new files that don't have old counterparts
                    Console.WriteLine("=== Creating additional PROJECTS to WIP mappings ===");
                    foreach (var newFile in newPrefixFiles)
                    {
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(newFile);
                        string extension = System.IO.Path.GetExtension(newFile);

                        // Create the expected old file name
                        string oldFileName = fileName.Replace(newPrefix, oldPrefix, StringComparison.OrdinalIgnoreCase);

                        // Create PROJECTS path (old path)
                        string projectsPath = modelPath.Replace("\\WIP\\", "\\PROJECTS\\");
                        string oldFilePath = System.IO.Path.Combine(projectsPath, oldFileName + extension);

                        // Create WIP path (new path)
                        string newFilePath = System.IO.Path.Combine(modelPath, fileName + extension);

                        // Only add if not already mapped
                        if (!fileMapping.ContainsKey(oldFilePath))
                        {
                            fileMapping[oldFilePath] = newFilePath;
                            string fileType = extension.ToLower() == ".iam" ? "assembly" : "part";
                            Console.WriteLine($"Created additional {fileType} mapping (PROJECTS->WIP): {oldFileName + extension} -> {System.IO.Path.GetFileName(newFilePath)}");
                        }
                    }
                    Console.WriteLine($"=== Additional PROJECTS to WIP mappings created: {fileMapping.Count} ===");
                    return fileMapping; // Return early since we've created all mappings
                }

                // Special case: Create mappings from PROJECTS path to WIP path
                Console.WriteLine("=== Creating PROJECTS to WIP path mappings ===");
                foreach (var newFile in newPrefixFiles)
                {
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(newFile);
                    string extension = System.IO.Path.GetExtension(newFile);

                    // Create the expected old file name
                    string oldFileName = fileName.Replace(newPrefix, oldPrefix, StringComparison.OrdinalIgnoreCase);

                    // Create PROJECTS path (old path) - handle both WIP and PROJECTS in the path
                    string projectsPath = modelPath.Replace("\\WIP\\", "\\PROJECTS\\");
                    string oldFilePath = System.IO.Path.Combine(projectsPath, oldFileName + extension);

                    // Create WIP path (new path)
                    string newFilePath = System.IO.Path.Combine(modelPath, fileName + extension);

                    // Create mapping from PROJECTS path to WIP path
                    fileMapping[oldFilePath] = newFilePath;
                    string fileType = extension.ToLower() == ".iam" ? "assembly" : "part";
                    Console.WriteLine($"Created {fileType} mapping (PROJECTS->WIP): {oldFileName + extension} -> {System.IO.Path.GetFileName(newFilePath)}");
                }
                Console.WriteLine($"=== PROJECTS to WIP mappings created: {fileMapping.Count} ===");
                return fileMapping; // Return early since we've created all mappings
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error building file mapping: {ex.Message}");
            }

            return fileMapping;
        }

        /// <summary>
        /// Validates that all required paths exist and are accessible
        /// </summary>
        private bool ValidatePaths(string drawingsPath, string modelPath, string projectPath, DrawingUpdateResult result)
        {
            try
            {
                // Validate drawings path
                if (string.IsNullOrWhiteSpace(drawingsPath))
                {
                    result.FailedDrawings.Add("Drawings path is required and cannot be empty");
                    return false;
                }

                if (!System.IO.Directory.Exists(drawingsPath))
                {
                    result.FailedDrawings.Add($"Drawings directory not found: {drawingsPath}");
                    return false;
                }

                // Validate model path
                if (string.IsNullOrWhiteSpace(modelPath))
                {
                    result.FailedDrawings.Add("Model path is required and cannot be empty");
                    return false;
                }

                if (!System.IO.Directory.Exists(modelPath))
                {
                    result.FailedDrawings.Add($"Model directory not found: {modelPath}");
                    return false;
                }

                // Validate project path (optional but if provided, should exist)
                if (!string.IsNullOrWhiteSpace(projectPath) && !System.IO.Directory.Exists(projectPath))
                {
                    result.FailedDrawings.Add($"Project directory not found: {projectPath}");
                    return false;
                }

                // Check if directories are accessible
                try
                {
                    System.IO.Directory.GetFiles(drawingsPath, "*", System.IO.SearchOption.TopDirectoryOnly);
                    System.IO.Directory.GetFiles(modelPath, "*", System.IO.SearchOption.TopDirectoryOnly);
                    if (!string.IsNullOrWhiteSpace(projectPath))
                    {
                        System.IO.Directory.GetFiles(projectPath, "*", System.IO.SearchOption.TopDirectoryOnly);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    result.FailedDrawings.Add($"Access denied to directory: {ex.Message}");
                    return false;
                }
                catch (Exception ex)
                {
                    result.FailedDrawings.Add($"Error accessing directory: {ex.Message}");
                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                result.FailedDrawings.Add($"Path validation error: {ex.Message}");
                return false;
            }
        }




        /// <summary>
        /// Updates assembly references in drawing files to match renamed assemblies in the model folder
        /// </summary>


        /// <summary>
        /// Retrieves all assembly (.iam) files from the specified folder (non-recursive).
        /// </summary>
        private List<string> RetrieveAssembliesFromPath(string folderPath)
        {
            if (!Directory.Exists(folderPath))
                return new List<string>();

            return Directory.GetFiles(folderPath, "*.iam", SearchOption.TopDirectoryOnly).ToList();
        }

        /// <summary>
        /// Builds a mapping of old full file paths to new full file paths by replacing the old prefix with the new prefix.
        /// Considers both part and assembly files (recursive).
        /// </summary>
        private Dictionary<string, string> BuildRenameMapping(string modelFolder, string oldPrefix, string newPrefix)
        {
            var mapping = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            if (!Directory.Exists(modelFolder))
                return mapping;

            var files = Directory.GetFiles(modelFolder, "*.*", SearchOption.AllDirectories)
                .Where(f => f.EndsWith(".iam", StringComparison.OrdinalIgnoreCase) ||
                            f.EndsWith(".ipt", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var filePath in files)
            {
                var fileName = System.IO.Path.GetFileName(filePath);
                if (fileName.StartsWith(oldPrefix, StringComparison.OrdinalIgnoreCase))
                {
                    var newName = fileName.Replace(oldPrefix, newPrefix, StringComparison.OrdinalIgnoreCase);
                    var newFullPath = System.IO.Path.Combine(System.IO.Path.GetDirectoryName(filePath)!, newName);

                    if (!mapping.ContainsKey(filePath))
                        mapping.Add(filePath, newFullPath);
                }
            }

            return mapping;
        }

        /// <summary>
        /// Deletes files from the file system while handling errors.
        /// Returns summary result.
        /// </summary>
        private FileDeletionResultNew DeleteFilesNew(List<string> filesToDelete)
        {
            var result = new FileDeletionResultNew();

            foreach (var file in filesToDelete)
            {
                try
                {
                    if (System.IO.File.Exists(file))
                    {
                        System.IO.File.Delete(file);
                        result.DeletedFiles.Add(file);
                    }
                    else
                    {
                        result.NotFoundFiles.Add(file);
                    }
                }
                catch (UnauthorizedAccessException ex)
                {
                    Console.WriteLine($"Access denied deleting file: {file} - {ex.Message}");
                    result.AccessDeniedFiles.Add(file);
                }
                catch (IOException ex)
                {
                    Console.WriteLine($"IO error deleting file: {file} - {ex.Message}");
                    result.FailedFiles.Add(file);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error deleting file: {file} - {ex.Message}");
                    result.FailedFiles.Add(file);
                }
            }

            return result;
        }

        /// <summary>
        /// Class to hold information about deleted and failed files during cleanup.
        /// </summary>
        public class FileDeletionResultNew
        {
            public List<string> DeletedFiles { get; set; } = new();
            public List<string> FailedFiles { get; set; } = new();
            public List<string> NotFoundFiles { get; set; } = new();
            public List<string> AccessDeniedFiles { get; set; } = new();
        }

        /// <summary>
        /// Cleans up Inventor application to release COM resources.
        /// </summary>
        private void CleanupInventorAppNew()
        {
            if (_inventorApp != null)
            {
                try
                {
                    while (_inventorApp.Documents.Count > 0)
                    {
                        try
                        {
                            _inventorApp.Documents[1].Close(false);
                        }
                        catch { /* ignore individual document cleanup errors */ }
                    }
                    _inventorApp.Quit();
                    Marshal.ReleaseComObject(_inventorApp);
                }
                catch { /* ignore cleanup errors */ }
                _inventorApp = null;
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }


        /// <summary>
        /// Result class for file deletion operations
        /// </summary>
        public class FileDeletionResult
        {
            public List<string> SuccessfullyDeleted { get; set; } = new();
            public List<string> FailedToDelete { get; set; } = new();
            public List<string> NotFound { get; set; } = new();
            public List<string> AccessDenied { get; set; } = new();
            public int TotalFiles { get; set; }
            public int DeletedCount { get; set; }
            public int FailedCount { get; set; }
            public int NotFoundCount { get; set; }
            public int AccessDeniedCount { get; set; }
        }

        /// <summary>
        /// Safely deletes a list of files and returns detailed results
        /// </summary>
        /// <param name="filePaths">List of file paths to delete</param>
        /// <returns>Detailed result of the deletion operation</returns>
        public FileDeletionResult DeleteFiles(List<string> filePaths)
        {
            var result = new FileDeletionResult
            {
                TotalFiles = filePaths.Count
            };

            Console.WriteLine($"=== Starting File Deletion Operation ===");
            Console.WriteLine($"Total files to delete: {filePaths.Count}");

            foreach (string filePath in filePaths)
            {
                try
                {
                    // Check if file exists
                    if (!System.IO.File.Exists(filePath))
                    {
                        Console.WriteLine($"File not found: {filePath}");
                        result.NotFound.Add(filePath);
                        result.NotFoundCount++;
                        continue;
                    }

                    // Check if file is locked/accessible
                    if (IsFileLocked(filePath))
                    {
                        Console.WriteLine($"File is locked, cannot delete: {filePath}");
                        result.AccessDenied.Add(filePath);
                        result.AccessDeniedCount++;
                        continue;
                    }

                    // Attempt to delete the file
                    System.IO.File.Delete(filePath);
                    Console.WriteLine($"Successfully deleted: {filePath}");
                    result.SuccessfullyDeleted.Add(filePath);
                    result.DeletedCount++;
                }
                catch (UnauthorizedAccessException ex)
                {
                    Console.WriteLine($"Access denied when deleting {filePath}: {ex.Message}");
                    result.AccessDenied.Add(filePath);
                    result.AccessDeniedCount++;
                }
                catch (System.IO.IOException ex)
                {
                    Console.WriteLine($"IO error when deleting {filePath}: {ex.Message}");
                    result.FailedToDelete.Add(filePath);
                    result.FailedCount++;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error when deleting {filePath}: {ex.Message}");
                    result.FailedToDelete.Add(filePath);
                    result.FailedCount++;
                }
            }

            Console.WriteLine($"=== File Deletion Operation Completed ===");
            Console.WriteLine($"Successfully deleted: {result.DeletedCount}");
            Console.WriteLine($"Failed to delete: {result.FailedCount}");
            Console.WriteLine($"Not found: {result.NotFoundCount}");
            Console.WriteLine($"Access denied: {result.AccessDeniedCount}");

            return result;
        }
    }
}