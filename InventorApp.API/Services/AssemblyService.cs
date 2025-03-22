using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using Inventor;
using InventorApp.API.Models;

namespace InventorApp.API.Services
{
    public class AssemblyService
    {
        private Inventor.Application? _inventorApp;
        private bool _isAssemblyOpen = false;

        public bool IsAssemblyOpen => _isAssemblyOpen;

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
                }
            }
        }

        public void ChangeParameters(string partFilePath, List<Dictionary<string, object>> parameters)
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

                Document partDoc = _inventorApp.Documents.Open(partFilePath);
                PartDocument part = (PartDocument)partDoc;
                Parameters paramList = part.ComponentDefinition.Parameters;

                foreach (var param in parameters)
                {
                    if (param.TryGetValue("parameterName", out var paramNameObj) && paramNameObj != null && param.TryGetValue("newValue", out var newValueObj))
                    {
                        string paramName = paramNameObj.ToString()!;
                        if (double.TryParse(newValueObj.ToString(), out double newValue))
                        {
                            paramList[paramName].Expression = $"{newValue} mm";
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

                partDoc.Save();
                partDoc.Close();
                Console.WriteLine($"Parameters updated successfully in {partFilePath}");
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error changing parameters: {e.Message}");
                throw;
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
                foreach (var action in suppressActions)
                {
                    string assemblyPath = System.IO.Path.Combine("D:\\Project_task\\Projects\\TRANSFORMER\\WIP\\PC0300949_01_01\\MODEL", action.AssemblyFilePath);

                    foreach (var component in action.Components)
                    {
                        SuppressComponent(assemblyPath, component, action.Suppress);
                    }
                }
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error suppressing multiple components: {e.Message}");
                throw;
            }
        }


        public void SuppressComponent(string filePath, string componentName, bool suppress)
        {
            try
            {
                // Ensure Inventor is running
                if (_inventorApp == null)
                {
                    Type? inventorType = Type.GetTypeFromProgID("Inventor.Application");
                    if (inventorType == null) throw new InvalidOperationException("Autodesk Inventor is not installed or registered.");

                    _inventorApp = (Inventor.Application)Activator.CreateInstance(inventorType)!;
                    _inventorApp.Visible = true;
                }

                // Open the document
                Inventor.Document doc = _inventorApp.Documents.Open(filePath, true);

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

                // Save and close
                doc.Save();
                doc.Close();
            }
            catch (Exception e)
            {
                Console.Error.WriteLine($"Error suppressing component: {e.Message}");
                throw;
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



        internal bool UpdateIPropertiesForAssemblies(List<Dictionary<string, object>> assemblyUpdates)
        {
            throw new NotImplementedException();
        }

        internal bool UpdateIPropertiesForAllFiles(string directoryPath, Dictionary<string, string> properties)
        {
            throw new NotImplementedException();
        }
        internal bool UpdateIpartsAndIassemblies(List<AssemblyUpdate> assemblyUpdates)
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

                foreach (var update in assemblyUpdates)
                {
                    string assemblyFilePath = System.IO.Path.Combine("D:\\Project_task\\Projects\\TRANSFORMER\\WIP\\PC0300949_01_01\\MODEL", update.AssemblyFilePath);

                    AssemblyDocument assemblyDoc = (AssemblyDocument)_inventorApp.Documents.Open(assemblyFilePath);
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
#pragma warning disable CS8604 // Possible null reference argument.
                                                string factoryPath = System.IO.Path.Combine(
                                                    System.IO.Path.GetDirectoryName(docPath),
                                                    System.IO.Path.GetFileNameWithoutExtension(docPath).Split(':')[0] + ".ipt");
#pragma warning restore CS8604 // Possible null reference argument.

                                                // Open the factory document
                                                Document factoryDoc = _inventorApp.Documents.Open(factoryPath);
                                                PartDocument factoryPartDoc = (PartDocument)factoryDoc;

                                                // Create the new member file path
#pragma warning disable CS8604 // Possible null reference argument.
                                                string newMemberPath = System.IO.Path.Combine(
                                                    System.IO.Path.GetDirectoryName(factoryPath),
                                                    newComponent + ".ipt");
#pragma warning restore CS8604 // Possible null reference argument.

                                                // Replace the occurrence with the new member
                                                targetOccurrence.Replace(newMemberPath, false);

                                                // Close the factory document
                                                factoryDoc.Close();

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
#pragma warning disable CS8604 // Possible null reference argument.
                                        string newPath = System.IO.Path.Combine(
                                            System.IO.Path.GetDirectoryName(occurrence.ReferencedDocumentDescriptor.FullDocumentName),
                                            newComponent + ".ipt");
#pragma warning restore CS8604 // Possible null reference argument.

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
                                    // Instead of trying to modify the iAssembly directly, let's replace the occurrence
                                    // with the desired configuration
#pragma warning disable CS8604 // Possible null reference argument.
                                    string newFullPath = System.IO.Path.Combine(
                                        System.IO.Path.GetDirectoryName(occurrence.ReferencedDocumentDescriptor.FullDocumentName),
                                        newComponent + ".iam");
#pragma warning restore CS8604 // Possible null reference argument.

                                    if (System.IO.File.Exists(newFullPath))
                                    {
                                        // Replace the occurrence with the new configuration file
                                        occurrence.Replace(newFullPath, false);
                                        Console.WriteLine($"Successfully replaced iAssembly with {newComponent}");
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Could not find iAssembly configuration file: {newFullPath}");

                                        // Alternative approach - try to use the Document.ReplaceiAssemblyMember method if available
                                        try
                                        {
                                            Document doc = (Document)occurrence.ReferencedDocumentDescriptor;
                                            if (doc != null && doc is AssemblyDocument asmDoc)
                                            {
                                                // Try to use reflection to find and invoke the method
                                                var method = asmDoc.GetType().GetMethod("ReplaceiAssemblyMember");
                                                if (method != null)
                                                {
                                                    method.Invoke(asmDoc, new object[] { newComponent });
                                                    Console.WriteLine($"Successfully updated iAssembly using reflection");
                                                }
                                                else
                                                {
                                                    Console.WriteLine("ReplaceiAssemblyMember method not found");
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            Console.WriteLine($"Alternative approach failed: {ex.Message}");
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
                        _inventorApp.ActiveView.Update();
                        assemblyDoc.Save();
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine($"Error saving assembly: {e.Message}");
                    }
                    finally
                    {
                        assemblyDoc.Close();
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.Error.WriteLine($"Error updating iParts/iAssemblies: {ex.Message}");
                return false;
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

    }
}
