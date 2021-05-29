#region Copyright (c) 2011-2021 Technosoftware GmbH. All rights reserved
//-----------------------------------------------------------------------------
// Copyright (c) 2011-2021 Technosoftware GmbH. All rights reserved
// Web: https://technosoftware.com 
// 
// License: 
// 
// Permission is hereby granted, free of charge, to any person obtaining a copy
// of this software and associated documentation files (the "Software"), to deal
// in the Software without restriction, including without limitation the rights
// to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
// copies of the Software, and to permit persons to whom the Software is
// furnished to do so, subject to the following conditions:

// The above copyright notice and this permission notice shall be included in all
// copies or substantial portions of the Software.

// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
// IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
// FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
// AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
// LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
// OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
// SOFTWARE.
//
// SPDX-License-Identifier: MIT
//-----------------------------------------------------------------------------
#endregion Copyright (c) 2011-2021 Technosoftware GmbH. All rights reserved

#region	Using Directives
using System;
using System.Threading;
using System.Collections;
using System.Collections.Generic;
#endregion

namespace ServerPlugin
{

    /// <summary>
    /// OPC Server Configuration and IO Handling
    ///
    /// This C# based plugin for the OPC Server .NET Classic Edition shows a base 
    /// OPC 2.05a / 3.00 server implementation  
    /// At startup items with several data types and access rights are statically 
    /// defined. 
    /// The RefreshThread simulates signal changes for the items
    ///      SimulatedData.Ramp
    ///      SimulatedData.Random
    ///      SimulatedData.Sine
    /// and writes the changed values into the internal cache. The generic server
    /// read the item values from device as required through calling the function 
    /// ReadItem().
    /// Item values written by a client are written into the local buffer only.
    /// </summary>
    public class ClassicNodeManager : ClassicBaseNodeManager
    {
        #region Constants

        #region Data Access IDs

        const int PropertyIdCasingMaterial          = 5650;
        const int PropertyIdCasingHeight            = 5651;
        const int PropertyIdCasingManufacturer      = 5652;

        #endregion

        #endregion

        #region Fields

        private static int itemHandle_ = 1;

        // Simulated Data Items
        private static MyItem myDynamicRampItem_;
        private static MyItem myDynamicSineItem_;
        private static MyItem myDynamicRandomItem_;

        #endregion

        #region Static Fields

        static private Thread myThread_;
        static private ManualResetEvent stopThread_;

        #endregion

        #region Signal State Data

        // DATA DEFINITIONS
        // Important: All data needs to be defined as STATIC.
        // This is important because this class is used in multiple instances.

        static private Dictionary<IntPtr, MyItem> items_ = new Dictionary<IntPtr, MyItem>();
        #endregion

        #region General Methods (not related to an OPC specification)

        #region  .NET API Generic Server Default Methods
        //---------------------------------------------------------------------
        //  .NET API Methods 
        // (Called by the generic server)
        //---------------------------------------------------------------------

        /// <summary>
        /// Gets the logging level to be used.
        /// </summary>
        /// <returns>
        ///     A LogLevel
        /// </returns>
        public override int OnGetLogLevel()
        {
            return (int)LogLevel.Info;
        }

        /// <summary>
        /// Gets the logging path to be used.
        /// </summary>
        /// <returns>
        ///     Path to be used for logging.
        /// </returns>
        public override string OnGetLogPath()
        {
            return "";
        }
		
        /// <summary>
        /// 	<para>
        ///         This method is called from the generic server at the startup; when the first
        ///         client connects or the service is started. All items supported by the server
        ///         need to be defined by calling the <see cref="AddItem">AddItem</see> or
        ///         <see cref="ClassicBaseNodeManager.AddAnalogItem">AddAnalogItem</see> callback method for each item.
        ///     </para>
        /// 	<para>The Item IDs are fully qualified names ( e.g. Dev1.Chn5.Temp ).</para>
        /// 	<para>
        ///         If <see cref="DaBrowseMode">DaBrowseMode.Generic</see> is set the generic
        ///         server part creates an approriate hierarchical address space. The sample code
        ///         defines the application item handle as the buffer array index. This handle is
        ///         passed in the calls from the generic server to identify the item. It should
        ///         allow quick access to the item definition / buffer. The handle may be
        ///         implemented differently depending on the application.
        ///     </para>
        /// 	<para>The branch separator character used in the fully qualified item name must
        ///     match the separator character defined in the OnGetDAServerParameters method.</para>
        /// </summary>
        /// <returns>A <see cref="StatusCodes"/> code with the result of the operation.</returns>
        public override int OnCreateServerItems()
        {
            #region Data Access address space

            CreateServerAddressSpace();

            #endregion

            // create a thread for simulating signal changes
            // in real application this thread reads from the device
            myThread_ = new Thread(RefreshThread) { Name = "Device Simulation", Priority = ThreadPriority.AboveNormal };
            myThread_.Start();

            return StatusCodes.Good;
        }

        public override void OnShutdownSignal()
        {
            //////////////////  TO-DO  /////////////////
            // close the device communication

            // terminate the simulation thread
            stopThread_ = new ManualResetEvent(false);
            stopThread_.WaitOne(5000, true);
            stopThread_.Close();
            stopThread_ = null;
        }

        public override int OnGetDaServerParameters(out int updatePeriod, out char branchDelimiter, out DaBrowseMode browseMode)
        {
            // Default Values
            updatePeriod = 100;                             // ms
            branchDelimiter = '.';
            browseMode = DaBrowseMode.Generic;            // browse the generic server address space
            return StatusCodes.Good;
        }

        #endregion

        #endregion

        #region Data Access related Methods

        #region  .NET API Generic Server Default Methods
        //---------------------------------------------------------------------
        //  .NET API Methods 
        // (Called by the generic server)
        //---------------------------------------------------------------------

        public override ClassicServerDefinition OnGetDaServerDefinition()
        {
            DaServer = new ClassicServerDefinition
            {
                ClsIdApp = "{FB2A8966-6B9B-4205-BDE3-18D87B80240B}",
                CompanyName = "Technosoftware GmbH",
                ClsIdServer = "{C08E1A10-A3DF-41F1-BA42-745B0004C8D0}",
                PrgIdServer = "OpcNetDaAe.DaSimpleSample",
                PrgIdCurrServer = "OpcNetDaAe.DaSimpleSample.90",
                ServerName = "OPC Server SDK .NET DA Simple Sample Server",
                CurrServerName = "OPC Server SDK .NET DA Simple Sample Server V9.0"
            };

            return DaServer;
        }

        /// <summary>
        /// Query the properties defined for the specified item
        /// </summary>
        /// <param name="deviceItemHandle">Generic Server device item handle</param>
        /// <param name="noProp">Number of properties returned</param>
        /// <param name="iDs">Array with the the property ID number</param>
        /// <returns>A <see cref="StatusCodes" /> code with the result of the operation. 
        ///  StatusCodes.Bad if the item has no custom properties.</returns>
        public override int OnQueryProperties(
            IntPtr deviceItemHandle,
            out int noProp,
            out int[] iDs)
        {
            MyItem item;
            if (items_.TryGetValue(deviceItemHandle, out item))
            {
                if (item.ItemProperties != null)
                {
                    // item has  custom properties
                    noProp = item.ItemProperties.Length;
                    iDs = new int[noProp];
                    for (int i = 0; i < noProp; ++i)
                    {
                        iDs[i] = item.ItemProperties[i].PropertyId;
                    }
                    return StatusCodes.Good;
                }
            }
            noProp = 0;
            iDs = null;
            return StatusCodes.Bad;
        }

        /// <summary>
        /// Returns the values of the requested custom properties of the requested item. This
        /// method is not called for the OPC standard properties 1..8. These are handled in the
        /// generic server.
        /// </summary>
        /// <returns>HRESULT success/error code. Bad if the item has no custom properties.</returns>
        /// <param name="deviceItemHandle">Generic Server device item handle</param>
        /// <param name="propertyId">ID of the property</param>
        /// <param name="propertyValue">Property value</param>
        public override int OnGetPropertyValue(IntPtr deviceItemHandle, int propertyId, out object propertyValue)
        {
            MyItem item;
            if (items_.TryGetValue(deviceItemHandle, out item))
            {
                if (item.ItemProperties != null)
                {

                    int numProp = item.ItemProperties.Length;
                    for (int i = 0; i < numProp; ++i)
                    {
                        if (item.ItemProperties[i].PropertyId == propertyId)
                        {
                            propertyValue = item.ItemProperties[i].PropertyValue;
                            return StatusCodes.Good;
                        }
                    }
                }
            }
            // Item property is not available
            propertyValue = null;
            return StatusCodes.BadInvalidPropertyId;
        }

        /// <summary>
        /// 	<para>This method is called when a client executes a 'write' server call. The items
        ///     specified in the DaDeviceItemValue array need to be written to the device.</para>
        /// 	<para>The cache is updated in the generic server after returning from the
        ///     customization WiteItems method. Items with write error are not updated in the
        ///     cache.</para>
        /// </summary>
        /// <returns>A <see cref="StatusCodes"/> code with the result of the operation.</returns>
        /// <param name="values">Object with handle, value, quality, timestamp</param>
        /// <param name="errors">Array with HRESULT success/error codes on return.</param>
        public override int OnWriteItems(DaDeviceItemValue[] values, out int[] errors)
        {
            errors = new int[values.Length];                            // result array
            for (int i = 0; i < values.Length; ++i)                     // init to Good
                errors[i] = StatusCodes.Good;

            // TO-DO: write the new values to the device
            foreach (DaDeviceItemValue t in values)
            {
                MyItem item;
                if (items_.TryGetValue(t.DeviceItemHandle, out item))
                {
                    // Only if there is a Value specified write the value into buffer
                    if (t.Value != null)
                        item.Value = t.Value;
                    if (t.QualitySpecified)
                        item.Quality = new DaQuality(t.Quality);
                    if (t.TimestampSpecified)
                        item.Timestamp = t.Timestamp;
                }
            }
            return StatusCodes.Good;
        }

        #endregion

        #endregion

        #region Create Data Access Sample Variants

        /// <summary>
        /// Utility function to create the sample values for the server items.
        /// </summary>
        /// <param name="itemType"></param>
        /// <param name="isArray"></param>
        /// <param name="itemValue"></param>
        internal void CreateSampleVariant(
            Type itemType,
            bool isArray,
            out object itemValue)
        {
            #region Data Access address space

            var rand = new Random();

            if (isArray)
            {
                //
                // Array Type
                //
                string bstr = null;
                if (itemType == typeof(string[]))
                {
                    bstr = "This is string #";
                }

                var itemList = new ArrayList();

                for (int i = 0; i < 4; i++)
                {
                    if (itemType == typeof(bool))
                    {
                        itemList.Add(i % 2 == 1);
                    }
                    else if (itemType == typeof(sbyte))
                    {
                        var d = new byte[1];
                        rand.NextBytes(d);
                        itemList.Add((sbyte)d[0]);
                    }
                    else if (itemType == typeof(short))
                    {
                        var d = new byte[1];
                        rand.NextBytes(d);
                        itemList.Add((short)d[0]);
                    }
                    else if (itemType == typeof(int))
                    {
                        itemList.Add(rand.Next() * 100);
                    }
                    else if (itemType == typeof(Int64))
                    {
                        itemList.Add((Int64)rand.Next() * 100);
                    }
                    else if (itemType == typeof(UInt64))
                    {
                        itemList.Add((UInt64)rand.Next() * 100);
                    }
                    else if (itemType == typeof(byte))
                    {
                        var d = new byte[1];
                        rand.NextBytes(d);
                        itemList.Add(d[0]);
                    }
                    else if (itemType == typeof(ushort))
                    {
                        var d = new byte[1];
                        rand.NextBytes(d);
                        itemList.Add((ushort)d[0]);
                    }
                    else if (itemType == typeof(uint))
                    {
                        itemList.Add((uint)rand.NextDouble() * 100);
                    }
                    else if (itemType == typeof(float))
                    {
                        itemList.Add((float)rand.NextDouble() * 100);
                    }
                    else if (itemType == typeof(double))
                    {
                        itemList.Add(rand.NextDouble() * 100);
                    }
                    else if (itemType == typeof(DateTime))
                    {
                        itemList.Add(new DateTime(rand.Next()));
                    }
                    else if (itemType == typeof(string))
                    {
                        itemList.Add(bstr + i);
                    }
                    else
                        itemList.Add(null);    // not supported type
                }
                itemValue = itemList.ToArray(itemType);
            }
            else
            {
                //
                // Simple Type
                //
                if (itemType == typeof(sbyte)) itemValue = (sbyte)76;
                else if (itemType == typeof(byte)) itemValue = (byte)23;
                else if (itemType == typeof(short)) itemValue = (short)345;
                else if (itemType == typeof(ushort)) itemValue = (ushort)39874;
                else if (itemType == typeof(int)) itemValue = 20196;
                else if (itemType == typeof(Int64)) itemValue = Int64.MinValue;
                else if (itemType == typeof(UInt64)) itemValue = UInt64.MaxValue;
                else if (itemType == typeof(uint)) itemValue = (uint)4230498;
                else if (itemType == typeof(float)) itemValue = (float)8.123242;
                else if (itemType == typeof(double)) itemValue = 83289.48243;
                else if (itemType == typeof(DateTime)) itemValue = new DateTime(1900, 1, 1, 12, 0, 0);
                else if (itemType == typeof(bool)) itemValue = false;
                else if (itemType == typeof(string)) itemValue = "-- It's a nice day --";
                else itemValue = null;
            }

            #endregion

        }

        internal class StructItemIDs
        {
            public StructItemIDs(string itemId, Type type)
            { ItemId = itemId; ItemType = type; }

            public string ItemId { get; set; }

            public Type ItemType { get; set; }
        }

        internal class StructIoTypes
        {
            public StructIoTypes(string branch, DaAccessRights accessRights)
            { Branch = branch; AccessRights = accessRights; }

            public string Branch { get; set; }

            public DaAccessRights AccessRights { get; set; }
        }

        /// <summary>
        /// Create all items supported by this server
        /// </summary>
        /// <returns></returns>
        private void CreateServerAddressSpace()
        {
            var arrayItems =
                new[]{
									   new StructItemIDs( "Short",        typeof(short)  ),
									   new StructItemIDs( "Integer",      typeof(int)    ),
									   new StructItemIDs( "Int64",        typeof(Int64)  ),
									   new StructItemIDs( "UInt64",       typeof(UInt64)  ),
									   new StructItemIDs( "SingleFloat",  typeof(float)  ),
									   new StructItemIDs( "DoubleFloat",  typeof(double) ),
									   new StructItemIDs( "String",       typeof(string) ),
									   new StructItemIDs( "Byte",         typeof(byte)   ),
									   new StructItemIDs( "Character",    typeof(sbyte)  ),
									   new StructItemIDs( "Word",         typeof(ushort) ),
									   new StructItemIDs( "DoubleWord",   typeof(uint)   ),
									   new StructItemIDs( "Boolean",      typeof(bool)   ),
									   new StructItemIDs( "DateTime",     typeof(DateTime)   ),
									   new StructItemIDs( null,           null        ) };

            var ioTypes =
                new[] {
										new StructIoTypes( "In", DaAccessRights.Readable      ),
										new StructIoTypes( "Out",   DaAccessRights.Writable      ),
										new StructIoTypes( "InOut", DaAccessRights.ReadWritable  ),
										new StructIoTypes( null, 0                    ) };
            int z;
            MyItem myItem;

                //------------------------------------------------------------
                // SimpleTypes In/Out/InOut
                //------------------------------------------------------------
            int i = 0;
                while (ioTypes[i].Branch != null)
                {
                    z = 0;
                    while (arrayItems[z].ItemId != null)
                    {
                        object initialItemValue;
                        CreateSampleVariant(arrayItems[z].ItemType, false, out initialItemValue);

                    var itemId = "CTT.SimpleTypes.";
                    itemId += ioTypes[i].Branch;


                    myItem = new MyItem(itemId + "." + arrayItems[z].ItemId, initialItemValue);

                    AddItem(itemId + "." + arrayItems[z].ItemId,
                            ioTypes[i].AccessRights, initialItemValue, out myItem.DeviceItemHandle);
                    items_.Add(myItem.DeviceItemHandle, myItem);
                        z++;
                    }
                    i++;
                }

                //------------------------------------------------------------
                // Arrays In/Out/InOut
                //------------------------------------------------------------
                i = 0;
                while (ioTypes[i].Branch != null)
                {
                    z = 0;
                    while (arrayItems[z].ItemId != null)
                    {
                        object initialItemValue;
                        CreateSampleVariant(arrayItems[z].ItemType, true, out initialItemValue);

                    var itemId = "CTT.ArrayTypes.";
                    itemId += ioTypes[i].Branch;

                    myItem = new MyItem(itemId + "." + arrayItems[z].ItemId + "[]", initialItemValue);

                    AddItem(itemId + "." + arrayItems[z].ItemId + "[]",
                            ioTypes[i].AccessRights, initialItemValue, out myItem.DeviceItemHandle);
                    items_.Add(myItem.DeviceItemHandle, myItem);
                        z++;
                    }
                    i++;
                }


                // SimulatedData/Ramp
                {
                const int itemValue = 0; // canonical data type

                myItem = new MyItem("SimulatedData.Ramp", itemValue);

                myDynamicRampItem_ = myItem;

                AddItem("SimulatedData.Ramp",
                        DaAccessRights.Readable, itemValue, out myItem.DeviceItemHandle);
                items_.Add(myItem.DeviceItemHandle, myItem);
                }

                // SimulatedData/Sine
                {
                const double itemValue = 0.0; // canonical data type

                myItem = new MyItem("SimulatedData.Sine", itemValue);

                myDynamicSineItem_ = myItem;

                AddItem("SimulatedData.Sine",
                        DaAccessRights.Readable, itemValue, out myItem.DeviceItemHandle);
                items_.Add(myItem.DeviceItemHandle, myItem);
                }

                // SimulatedData/Random
                {
                const int itemValue = 0; // canonical data type

                myItem = new MyItem("SimulatedData.Random", itemValue);

                myDynamicRandomItem_ = myItem;

                AddItem(myItem.ItemName, DaAccessRights.Readable, itemValue, out myItem.DeviceItemHandle);
                items_.Add(myItem.DeviceItemHandle, myItem);
                }

                // SpecialItems/WithAnalogEUInfo
            var itemPropertiesAnalog = new MyItemProperty[2];
            itemPropertiesAnalog[0] = new MyItemProperty(DaProperty.LowEu.Code, 40.86);
            itemPropertiesAnalog[1] = new MyItemProperty(DaProperty.HighEu.Code, 92.67);

            myItem = new MyItem("SpecialItems.WithAnalogEUInfo", 20.56, itemPropertiesAnalog);

            AddAnalogItem(myItem.ItemName,
                          DaAccessRights.ReadWritable, myItem.Value, 40.86, 92.67, out myItem.DeviceItemHandle);
            items_.Add(myItem.DeviceItemHandle, myItem);

                // SpecialItems/WithAnalogEUInfo
            itemPropertiesAnalog[0] = new MyItemProperty(DaProperty.LowEu.Code, 12.50);
            itemPropertiesAnalog[1] = new MyItemProperty(DaProperty.HighEu.Code, 27.90);

            myItem = new MyItem("SpecialItems.WithAnalogEUInfo2", 21.00, itemPropertiesAnalog);

            AddAnalogItem(myItem.ItemName,
                          DaAccessRights.ReadWritable, myItem.Value, 12.50, 27.90, out myItem.DeviceItemHandle);
            items_.Add(myItem.DeviceItemHandle, myItem);

                // Add Custom Property Definitions to the generic server
            AddProperty(PropertyIdCasingHeight, "Casing Height", 25.34);
            AddProperty(PropertyIdCasingMaterial, "Casing Material", "Aluminum");
            AddProperty(PropertyIdCasingManufacturer, "Casing Manufacturer", "CBM");
            AddProperty(102, "High EU", 45.86);
            AddProperty(103, "Low EU", 35.86);

                // Create custom item properties for the item
            var itemProperties = new MyItemProperty[3];
            itemProperties[0] = new MyItemProperty(PropertyIdCasingHeight, 25.45);
            itemProperties[1] = new MyItemProperty(PropertyIdCasingMaterial, "Aluminum");
            itemProperties[2] = new MyItemProperty(PropertyIdCasingManufacturer, "CBM");

            myItem = new MyItem("SpecialItems.WithVendorSpecificProperties", 1111, itemProperties);

            AddItem(myItem.ItemName,
                    DaAccessRights.ReadWritable, myItem.Value, out myItem.DeviceItemHandle);
            items_.Add(myItem.DeviceItemHandle, myItem);
        }
        #endregion

        #region Refresh Thread

        // This method simulates item value changes.
        void RefreshThread()
        {
            double count = 0;
            int ramp = 0;
            var rand = new Random();

            // Update all used items once
            foreach (MyItem item in items_.Values)
            {
                SetItemValue(item.DeviceItemHandle, item.Value, DaQuality.Good.Code, DateTime.Now);
            }

            for (; ; )   // forever thread loop
            {
                //int numclientHandles;
                //IntPtr[] clientHandles;
                //String[] clientNames;

                //GetClients(out numclientHandles, out clientHandles, out clientNames);
                //if (numclientHandles > 0)
                //{
                //    foreach (var client in clientHandles)
                //    {
                //        int numGroupHandles;
                //        IntPtr[] groupHandles;
                //        String[] groupNames;
                //        GetGroups(client, out numGroupHandles, out groupHandles, out groupNames);
                //        if (numGroupHandles > 0)
                //        {
                //            foreach (var group in groupHandles)
                //            {
                //                DaGroupState groupInfo;
                //                GetGroupState(group, out groupInfo);
                //            }
                //        }
                //    }
                //}
                
                // Sample how to get a list of active items (used at least by one client)
                //int numHandles;
                //IntPtr[] itemHandles;
                //GetActiveItems(out numHandles, out itemHandles);

                count++;
                ramp++;
                myDynamicRampItem_.Value = ramp;
                // update server cache for this item
                SetItemValue(myDynamicRampItem_.DeviceItemHandle, myDynamicRampItem_.Value,
                   DaQuality.Good.Code, DateTime.Now);

                myDynamicRandomItem_.Value = rand.Next();
                SetItemValue(myDynamicRandomItem_.DeviceItemHandle, myDynamicRandomItem_.Value,
                   DaQuality.Good.Code, DateTime.Now);

                myDynamicSineItem_.Value = Math.Sin((count % 40) * 0.1570796327);
                SetItemValue(myDynamicSineItem_.DeviceItemHandle, myDynamicSineItem_.Value,
                   DaQuality.Good.Code, DateTime.Now);

                Thread.Sleep(1000);    // ms

                if (stopThread_ != null)
                {
                    stopThread_.Set();
                    return;               // terminate the thread
                }
            }
        }
        #endregion

    }

    #region My Item Property Class

    /// <summary>
    /// My Item Property Implementation
    /// </summary>
    public class MyItemProperty
    {
        #region Constructors, Destructor, Initialization

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="propertyId">ID of the property</param>
        /// <param name="propertyValue">Value of the property</param>
        public MyItemProperty(int propertyId, object propertyValue)
        {
            PropertyId = propertyId;
            PropertyValue = propertyValue;
        }

        #endregion

        #region Properties

        /// <summary>
        /// ID of the property
        /// </summary>
        public int PropertyId { get; private set; }

        /// <summary>
        /// Value of the property
        /// </summary>
        public object PropertyValue { get; private set; }

        #endregion

    }

    #endregion

    #region My Item Class

    /// <summary>
    /// My Item Implementation
    /// </summary>
    class MyItem
    {

        #region Constructors, Destructor, Initialization

        public MyItem(
                        string itemName,
                        object initValue)
        {
            ItemName = itemName;
            Value = initValue;
            Quality = DaQuality.Good;
            Timestamp = DateTime.UtcNow;
        }

        public MyItem(
                string itemName,
                object initValue,
                MyItemProperty[] itemProperties)
        {
            ItemName = itemName;
            Value = initValue;
            Quality = DaQuality.Good;
            ItemProperties = itemProperties;
            Timestamp = DateTime.UtcNow;
        }

        #endregion

        #region Properties

        // Can be used to identify the item, not used in this example. You can use also other information like device
        // specific information (e.g. serial line, datablock and data number for PLC, ...
        public IntPtr DeviceItemHandle;
        public string ItemName { get; private set; }
        public object Value { get; set; }
        public DaQuality Quality { get;  set; }
        public DateTime Timestamp { get; set; }
        public MyItemProperty[] ItemProperties { get; private set; }

        #endregion
    }

    #endregion

}