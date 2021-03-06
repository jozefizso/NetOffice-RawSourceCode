﻿using System;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using NetOffice;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi.Events
{	
	#pragma warning disable
	
	#region SinkPoint Interface

	[SupportByVersion("MSForms", 2)]
    [InternalEntity(InternalEntityKind.ComEventInterface)]
    [ComImport, Guid("7B020EC7-AF6C-11CE-9F46-00AA00574A4F"), InterfaceType(ComInterfaceType.InterfaceIsIDispatch), TypeLibType((short)0x1010)]
	public interface TabStripEvents
	{
		[SupportByVersion("MSForms", 2)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [SinkArgument("data", typeof(MSFormsApi.DataObject))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("dragState", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmDragState))]
        [SinkArgument("effect", typeof(MSFormsApi.ReturnEffect))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(3)]
		void BeforeDragOver([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object dragState, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [SinkArgument("action", SinkArgumentType.Enum, typeof(MSFormsApi.Enums.fmAction))]
        [SinkArgument("data", typeof(MSFormsApi.DataObject))]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [SinkArgument("effect", typeof(MSFormsApi.ReturnEffect))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(4)]
		void BeforeDropOrPaste([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift);

		[SupportByVersion("MSForms", 2)]
		[PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(2)]
		void Change();

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-600)]
		void Click([In] object index);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("index", SinkArgumentType.Int32)]
        [SinkArgument("cancel", typeof(MSFormsApi.ReturnBoolean))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-601)]
		void DblClick([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("number", SinkArgumentType.Int32)]
        [SinkArgument("description", typeof(MSFormsApi.ReturnString))]
        [SinkArgument("sCode", SinkArgumentType.Int32)]
        [SinkArgument("source", SinkArgumentType.String)]
        [SinkArgument("helpFile", SinkArgumentType.String)]
        [SinkArgument("helpContext", SinkArgumentType.Int32)]
        [SinkArgument("cancelDisplay", typeof(MSFormsApi.ReturnBoolean))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-608)]
		void Error([In] object number, [In, MarshalAs(UnmanagedType.IDispatch)] object description, [In] object sCode, [In] object source, [In] object helpFile, [In] object helpContext, [In, MarshalAs(UnmanagedType.IDispatch)] object cancelDisplay);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("keyCode", typeof(MSFormsApi.ReturnInteger))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-602)]
		void KeyDown([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("keyAscii", typeof(MSFormsApi.ReturnInteger))]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-603)]
		void KeyPress([In, MarshalAs(UnmanagedType.IDispatch)] object keyAscii);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("keyCode", typeof(MSFormsApi.ReturnInteger))]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-604)]
		void KeyUp([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-605)]
		void MouseDown([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-606)]
		void MouseMove([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);

		[SupportByVersion("MSForms", 2)]
        [SinkArgument("button", SinkArgumentType.Int16)]
        [SinkArgument("shift", SinkArgumentType.Int16)]
        [SinkArgument("x", SinkArgumentType.Single)]
        [SinkArgument("y", SinkArgumentType.Single)]
        [PreserveSig, MethodImpl(MethodImplOptions.InternalCall, MethodCodeType = MethodCodeType.Runtime), DispId(-607)]
		void MouseUp([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y);
	}

    #endregion

    #region SinkHelper

    [InternalEntity(InternalEntityKind.SinkHelper)]
    [ComVisible(true), ClassInterface(ClassInterfaceType.None), TypeLibType(TypeLibTypeFlags.FHidden)]
	public class TabStripEvents_SinkHelper : SinkHelper, TabStripEvents
	{
		#region Static
		
		public static readonly string Id = "7B020EC7-AF6C-11CE-9F46-00AA00574A4F";
		
		#endregion	
		
		#region Ctor

		public TabStripEvents_SinkHelper(ICOMObject eventClass, IConnectionPoint connectPoint): base(eventClass)
		{
			SetupEventBinding(connectPoint);
		}
		
		#endregion

		#region TabStripEvents
		
		public void BeforeDragOver([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In] object dragState, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift)
        {
            if (!Validate("BeforeDragOver"))
            {
                Invoker.ReleaseParamsArray(index, cancel, data, x, y, dragState, effect, shift);
                return;
            }

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
			NetOffice.MSFormsApi.DataObject newData = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.DataObject>(EventClass, data, NetOffice.MSFormsApi.DataObject.LateBindingApiWrapperType);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			NetOffice.MSFormsApi.Enums.fmDragState newDragState = (NetOffice.MSFormsApi.Enums.fmDragState)dragState;
			NetOffice.MSFormsApi.ReturnEffect newEffect = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnEffect>(EventClass, effect, NetOffice.MSFormsApi.ReturnEffect.LateBindingApiWrapperType);
			Int16 newShift = ToInt16(shift);
			object[] paramsArray = new object[8];
			paramsArray[0] = newIndex;
			paramsArray[1] = newCancel;
			paramsArray[2] = newData;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			paramsArray[5] = newDragState;
			paramsArray[6] = newEffect;
			paramsArray[7] = newShift;
			EventBinding.RaiseCustomEvent("BeforeDragOver", ref paramsArray);
		}

		public void BeforeDropOrPaste([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel, [In] object action, [In, MarshalAs(UnmanagedType.IDispatch)] object data, [In] object x, [In] object y, [In, MarshalAs(UnmanagedType.IDispatch)] object effect, [In] object shift)
        {
            if (!Validate("BeforeDropOrPaste"))
            {
                Invoker.ReleaseParamsArray(index, cancel, action, data, x, y, effect, shift);
                return;
            }

			Int32 newIndex = Convert.ToInt32(index);
			NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
			NetOffice.MSFormsApi.Enums.fmAction newAction = (NetOffice.MSFormsApi.Enums.fmAction)action;
			NetOffice.MSFormsApi.DataObject newData = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.DataObject>(EventClass, data, NetOffice.MSFormsApi.DataObject.LateBindingApiWrapperType);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			NetOffice.MSFormsApi.ReturnEffect newEffect = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnEffect>(EventClass, effect, NetOffice.MSFormsApi.ReturnEffect.LateBindingApiWrapperType);
			Int16 newShift = ToInt16(shift);
			object[] paramsArray = new object[8];
			paramsArray[0] = newIndex;
			paramsArray[1] = newCancel;
			paramsArray[2] = newAction;
			paramsArray[3] = newData;
			paramsArray[4] = newX;
			paramsArray[5] = newY;
			paramsArray[6] = newEffect;
			paramsArray[7] = newShift;
			EventBinding.RaiseCustomEvent("BeforeDropOrPaste", ref paramsArray);
		}

		public void Change()
		{
            if (!Validate("Change"))
            {
                return;
            }

            object[] paramsArray = new object[0];
			EventBinding.RaiseCustomEvent("Change", ref paramsArray);
		}

		public void Click([In] object index)
		{
            if (!Validate("Click"))
            {
                return;
            }

            Int32 newIndex = ToInt32(index);
            object[] paramsArray = new object[1];
            paramsArray[0] = newIndex;
            EventBinding.RaiseCustomEvent("Click", ref paramsArray);
        }

		public void DblClick([In] object index, [In, MarshalAs(UnmanagedType.IDispatch)] object cancel)
		{
            if (!Validate("DblClick"))
            {
                Invoker.ReleaseParamsArray(index, cancel);
                return;
            }

            Int32 newIndex = ToInt32(index);
            NetOffice.MSFormsApi.ReturnBoolean newCancel = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancel, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
            object[] paramsArray = new object[2];
            paramsArray[0] = newIndex;
            paramsArray[1] = newCancel;
            EventBinding.RaiseCustomEvent("DblClick", ref paramsArray);
        }

        public void Error([In] object number, [In, MarshalAs(UnmanagedType.IDispatch)] object description, [In] object sCode, [In] object source, [In] object helpFile, [In] object helpContext, [In, MarshalAs(UnmanagedType.IDispatch)] object cancelDisplay)
		{
            if (!Validate("Error"))
            {
                Invoker.ReleaseParamsArray(number, description, sCode, source, helpFile, helpContext, cancelDisplay);
                return;
            }

			Int16 newNumber = ToInt16(number);
            NetOffice.MSFormsApi.ReturnString newDescription = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnString>(EventClass, description, NetOffice.MSFormsApi.ReturnString.LateBindingApiWrapperType);
            Int32 newSCode = ToInt32(sCode);
			string newSource = ToString(source);
			string newHelpFile = ToString(helpFile);
			Int32 newHelpContext = ToInt32(helpContext);
            NetOffice.MSFormsApi.ReturnBoolean newCancelDisplay = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnBoolean>(EventClass, cancelDisplay, NetOffice.MSFormsApi.ReturnBoolean.LateBindingApiWrapperType);
            object[] paramsArray = new object[7];
			paramsArray[0] = newNumber;
			paramsArray[1] = newDescription;
			paramsArray[2] = newSCode;
			paramsArray[3] = newSource;
			paramsArray[4] = newHelpFile;
			paramsArray[5] = newHelpContext;
			paramsArray[6] = newCancelDisplay;
			EventBinding.RaiseCustomEvent("Error", ref paramsArray);
		}

		public void KeyDown([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift)
        {
            if (!Validate("KeyDown"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

            NetOffice.MSFormsApi.ReturnInteger newKeyCode = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnInteger>(EventClass, keyCode, NetOffice.MSFormsApi.ReturnInteger.LateBindingApiWrapperType);
            Int16 newShift = ToInt16(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			EventBinding.RaiseCustomEvent("KeyDown", ref paramsArray);
		}

		public void KeyPress([In, MarshalAs(UnmanagedType.IDispatch)] object keyAscii)
        {
            if (!Validate("KeyPress"))
            {
                Invoker.ReleaseParamsArray(keyAscii);
                return;
            }

            NetOffice.MSFormsApi.ReturnInteger newKeyAscii = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnInteger>(EventClass, keyAscii, NetOffice.MSFormsApi.ReturnInteger.LateBindingApiWrapperType);
            object[] paramsArray = new object[1];
			paramsArray[0] = newKeyAscii;
			EventBinding.RaiseCustomEvent("KeyPress", ref paramsArray);
		}

		public void KeyUp([In, MarshalAs(UnmanagedType.IDispatch)] object keyCode, [In] object shift)
        {
            if (!Validate("KeyUp"))
            {
                Invoker.ReleaseParamsArray(keyCode, shift);
                return;
            }

            NetOffice.MSFormsApi.ReturnInteger newKeyCode = Factory.CreateKnownObjectFromComProxy<NetOffice.MSFormsApi.ReturnInteger>(EventClass, keyCode, NetOffice.MSFormsApi.ReturnInteger.LateBindingApiWrapperType);
            Int16 newShift = ToInt16(shift);
			object[] paramsArray = new object[2];
			paramsArray[0] = newKeyCode;
			paramsArray[1] = newShift;
			EventBinding.RaiseCustomEvent("KeyUp", ref paramsArray);
		}

		public void MouseDown([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y)
        {
            if (!Validate("MouseDown"))
            {
                Invoker.ReleaseParamsArray(index, button, shift, x, y);
                return;
            }

			Int32 newIndex = ToInt32(index);
			Int16 newButton = ToInt16(button);
			Int16 newShift = ToInt16(shift);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newIndex;
			paramsArray[1] = newButton;
			paramsArray[2] = newShift;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			EventBinding.RaiseCustomEvent("MouseDown", ref paramsArray);
		}

		public void MouseMove([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y)
        {
            if (!Validate("MouseMove"))
            {
                Invoker.ReleaseParamsArray(index, button, shift, x, y);
                return;
            }

			Int32 newIndex = ToInt32(index);
			Int16 newButton = ToInt16(button);
			Int16 newShift = ToInt16(shift);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newIndex;
			paramsArray[1] = newButton;
			paramsArray[2] = newShift;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			EventBinding.RaiseCustomEvent("MouseMove", ref paramsArray);
		}

		public void MouseUp([In] object index, [In] object button, [In] object shift, [In] object x, [In] object y)
		{
            if (!Validate("MouseUp"))
            {
                Invoker.ReleaseParamsArray(index, button, shift, x, y);
                return;
            }

            Int32 newIndex = ToInt32(index);
			Int16 newButton = ToInt16(button);
			Int16 newShift = ToInt16(shift);
			Single newX = ToSingle(x);
			Single newY = ToSingle(y);
			object[] paramsArray = new object[5];
			paramsArray[0] = newIndex;
			paramsArray[1] = newButton;
			paramsArray[2] = newShift;
			paramsArray[3] = newX;
			paramsArray[4] = newY;
			EventBinding.RaiseCustomEvent("MouseUp", ref paramsArray);
		}

		#endregion
	}
	
	#endregion
	
	#pragma warning restore
}