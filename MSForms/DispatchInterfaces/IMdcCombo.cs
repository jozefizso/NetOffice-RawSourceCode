using System;
using NetRuntimeSystem = System;
using System.ComponentModel;
using NetOffice.Attributes;

namespace NetOffice.MSFormsApi
{
	/// <summary>
	/// IMdcCombo
	/// </summary>
	[SyntaxBypass]
 	public class IMdcCombo_ : COMObject
	{
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IMdcCombo_(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMdcCombo_(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        /// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo_(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		/// <param name="parentObject">object there has created the proxy</param>
        /// <param name="comProxy">inner wrapped COM proxy</param>
        /// <param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo_(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo_(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}
		
		/// <param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo_(ICOMObject replacedObject) : base(replacedObject)
		{
		}

		/// <summary>
        /// Hidden stub .ctor
        /// </summary>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo_() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo_(string progId) : base(progId)
		{
		}
		
		#endregion

		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Column(object pvargColumn, object pvargIndex)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Column", pvargColumn, pvargIndex);
		}

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Column(object pvargColumn, object pvargIndex, object value)
		{
			Factory.ExecutePropertySet(this, "Column", pvargColumn, pvargIndex, value);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Alias for get_Column
		/// </summary>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2), Redirect("get_Column")]
		public object Column(object pvargColumn, object pvargIndex)
		{
			return get_Column(pvargColumn, pvargIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_Column(object pvargColumn)
		{
			return Factory.ExecuteVariantPropertyGet(this, "Column", pvargColumn);
		}

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_Column(object pvargColumn, object value)
		{
			Factory.ExecutePropertySet(this, "Column", pvargColumn, value);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Alias for get_Column
		/// </summary>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		[SupportByVersion("MSForms", 2), Redirect("get_Column")]
		public object Column(object pvargColumn)
		{
			return get_Column(pvargColumn);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_List(object pvargIndex, object pvargColumn)
		{
			return Factory.ExecuteVariantPropertyGet(this, "List", pvargIndex, pvargColumn);
		}

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="pvargColumn">optional object pvargColumn</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_List(object pvargIndex, object pvargColumn, object value)
		{
			Factory.ExecutePropertySet(this, "List", pvargIndex, pvargColumn, value);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Alias for get_List
		/// </summary>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		/// <param name="pvargColumn">optional object pvargColumn</param>
		[SupportByVersion("MSForms", 2), Redirect("get_List")]
		public object List(object pvargIndex, object pvargColumn)
		{
			return get_List(pvargIndex, pvargColumn);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object get_List(object pvargIndex)
		{
			return Factory.ExecuteVariantPropertyGet(this, "List", pvargIndex);
		}

        /// <summary>
        /// SupportByVersion MSForms 2
        /// Get/Set
        /// </summary>
        /// <param name="pvargIndex">optional object pvargIndex</param>
        /// <param name="value">optional object value</param>
        [SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public void set_List(object pvargIndex, object value)
		{
			Factory.ExecutePropertySet(this, "List", pvargIndex, value);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Alias for get_List
		/// </summary>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2), Redirect("get_List")]
		public object List(object pvargIndex)
		{
			return get_List(pvargIndex);
		}

		#endregion

		#region Methods

		#endregion
	}

	/// <summary>
	/// DispatchInterface IMdcCombo 
	/// SupportByVersion MSForms, 2
	/// </summary>
	[SupportByVersion("MSForms", 2)]
	[EntityType(EntityType.IsDispatchInterface), BaseType]
 	public class IMdcCombo : IMdcCombo_
	{
		#pragma warning disable

		#region Type Information

		/// <summary>
		/// Instance Type
		/// </summary>
		[EditorBrowsable(EditorBrowsableState.Advanced), Browsable(false), Category("NetOffice"), CoreOverridden]
		public override Type InstanceType
		{
			get
			{
				return LateBindingApiWrapperType;
			}
		}

        private static Type _type;

		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
        public static Type LateBindingApiWrapperType
        {
            get
            {
                if (null == _type)
                    _type = typeof(IMdcCombo);
                return _type;
            }
        }
        
        #endregion
        
		#region Ctor

		/// <param name="factory">current used factory core</param>
		/// <param name="parentObject">object there has created the proxy</param>
		/// <param name="proxyShare">proxy share instead if com proxy</param>
		public IMdcCombo(Core factory, ICOMObject parentObject, COMProxyShare proxyShare) : base(factory, parentObject, proxyShare)
		{
		}

		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
		public IMdcCombo(Core factory, ICOMObject parentObject, object comProxy) : base(factory, parentObject, comProxy)
		{
			
		}

        ///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo(ICOMObject parentObject, object comProxy) : base(parentObject, comProxy)
		{
		}
		
		///<param name="factory">current used factory core</param>
		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo(Core factory, ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(factory, parentObject, comProxy, comProxyType)
		{

		}

		///<param name="parentObject">object there has created the proxy</param>
        ///<param name="comProxy">inner wrapped COM proxy</param>
        ///<param name="comProxyType">Type of inner wrapped COM proxy"</param>
        [EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo(ICOMObject parentObject, object comProxy, NetRuntimeSystem.Type comProxyType) : base(parentObject, comProxy, comProxyType)
		{
		}
		
		///<param name="replacedObject">object to replaced. replacedObject are not usable after this action</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo(ICOMObject replacedObject) : base(replacedObject)
		{
		}
		
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo() : base()
		{
		}
		
		/// <param name="progId">registered progID</param>
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public IMdcCombo(string progId) : base(progId)
		{
		}
		
		#endregion
		
		#region Properties

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool AutoSize
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool AutoTab
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoTab");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoTab", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool AutoWordSelect
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "AutoWordSelect");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "AutoWordSelect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 BackColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BackColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BackColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmBackStyle BackStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmBackStyle>(this, "BackStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BackStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 BorderColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "BorderColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BorderColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmBorderStyle BorderStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmBorderStyle>(this, "BorderStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "BorderStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool BordersSuppress
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "BordersSuppress");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "BordersSuppress", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object BoundColumn
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "BoundColumn");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "BoundColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool CanPaste
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "CanPaste");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 ColumnCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ColumnCount");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnCount", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool ColumnHeads
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "ColumnHeads");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnHeads", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public string ColumnWidths
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "ColumnWidths");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ColumnWidths", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 CurTargetX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CurTargetX");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 CurTargetY
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CurTargetY");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 CurX
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "CurX");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "CurX", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmDropButtonStyle DropButtonStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmDropButtonStyle>(this, "DropButtonStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DropButtonStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool Enabled
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Enabled");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Enabled", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Font _Font_Reserved
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSFormsApi.Font>(this, "_Font_Reserved");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "_Font_Reserved", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[BaseResult]
		public NetOffice.MSFormsApi.Font Font
		{
			get
			{
				return Factory.ExecuteBaseReferencePropertyGet<NetOffice.MSFormsApi.Font>(this, "Font");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "Font", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontBold
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontBold");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontBold", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontItalic
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontItalic");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontItalic", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string FontName
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "FontName");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontName", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public float FontSize
		{
			get
			{
				return Factory.ExecuteFloatPropertyGet(this, "FontSize");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontSize", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontStrikethru
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontStrikethru");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontStrikethru", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool FontUnderline
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "FontUnderline");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontUnderline", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int16 FontWeight
		{
			get
			{
				return Factory.ExecuteInt16PropertyGet(this, "FontWeight");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "FontWeight", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 ForeColor
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ForeColor");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ForeColor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool HideSelection
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "HideSelection");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "HideSelection", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 LineCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "LineCount");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 ListCount
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ListCount");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// Unknown COM Proxy
		/// </summary>
		[SupportByVersion("MSForms", 2), ProxyResult]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ListCursor
		{
			get
			{
				return Factory.ExecuteReferencePropertyGet(this, "ListCursor");
			}
			set
			{
				Factory.ExecuteReferencePropertySet(this, "ListCursor", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object ListIndex
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ListIndex");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ListIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 ListRows
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "ListRows");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "ListRows", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmListStyle ListStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmListStyle>(this, "ListStyle");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ListStyle", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object ListWidth
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "ListWidth");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "ListWidth", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool Locked
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Locked");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Locked", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmMatchEntry MatchEntry
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMatchEntry>(this, "MatchEntry");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MatchEntry", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool MatchFound
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchFound");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool MatchRequired
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "MatchRequired");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MatchRequired", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public Int32 MaxLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "MaxLength");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "MaxLength", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2), NativeResult]
		public stdole.Picture MouseIcon
		{
			get
			{
				object[] paramsArray = null;
				object returnItem = Invoker.PropertyGet(this, "MouseIcon", paramsArray);
                return returnItem as stdole.Picture;
			}
			set
			{
				object[] paramsArray = Invoker.ValidateParamsArray(value);
				Invoker.PropertySet(this, "MouseIcon", paramsArray);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmMousePointer MousePointer
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmMousePointer>(this, "MousePointer");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "MousePointer", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public bool SelectionMargin
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "SelectionMargin");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelectionMargin", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 SelLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelLength");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelLength", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 SelStart
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "SelStart");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelStart", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public string SelText
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "SelText");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "SelText", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmShowDropButtonWhen ShowDropButtonWhen
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmShowDropButtonWhen>(this, "ShowDropButtonWhen");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "ShowDropButtonWhen", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmSpecialEffect SpecialEffect
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmSpecialEffect>(this, "SpecialEffect");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "SpecialEffect", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmStyle Style
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmStyle>(this, "Style");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "Style", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public string Text
		{
			get
			{
				return Factory.ExecuteStringPropertyGet(this, "Text");
			}
			set
			{
				Factory.ExecuteValuePropertySet(this, "Text", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmTextAlign TextAlign
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmTextAlign>(this, "TextAlign");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "TextAlign", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object TextColumn
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TextColumn");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TextColumn", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public Int32 TextLength
		{
			get
			{
				return Factory.ExecuteInt32PropertyGet(this, "TextLength");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object TopIndex
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "TopIndex");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "TopIndex", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public bool Valid
		{
			get
			{
				return Factory.ExecuteBoolPropertyGet(this, "Valid");
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public object Value
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Value");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Value", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object Column
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "Column");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "Column", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public object List
		{
			get
			{
				return Factory.ExecuteVariantPropertyGet(this, "List");
			}
			set
			{
				Factory.ExecuteVariantPropertySet(this, "List", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmIMEMode IMEMode
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmIMEMode>(this, "IMEMode");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "IMEMode", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmEnterFieldBehavior EnterFieldBehavior
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmEnterFieldBehavior>(this, "EnterFieldBehavior");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "EnterFieldBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get/Set
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public NetOffice.MSFormsApi.Enums.fmDragBehavior DragBehavior
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmDragBehavior>(this, "DragBehavior");
			}
			set
			{
				Factory.ExecuteEnumPropertySet(this, "DragBehavior", value);
			}
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// Get
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		[EditorBrowsable(EditorBrowsableState.Never), Browsable(false)]
		public NetOffice.MSFormsApi.Enums.fmDisplayStyle DisplayStyle
		{
			get
			{
				return Factory.ExecuteEnumPropertyGet<NetOffice.MSFormsApi.Enums.fmDisplayStyle>(this, "DisplayStyle");
			}
		}

		#endregion

		#region Methods

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pvargItem">optional object pvargItem</param>
		/// <param name="pvargIndex">optional object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		public void AddItem(object pvargItem, object pvargIndex)
		{
			 Factory.ExecuteMethod(this, "AddItem", pvargItem, pvargIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void AddItem()
		{
			 Factory.ExecuteMethod(this, "AddItem");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pvargItem">optional object pvargItem</param>
		[CustomMethod]
		[SupportByVersion("MSForms", 2)]
		public void AddItem(object pvargItem)
		{
			 Factory.ExecuteMethod(this, "AddItem", pvargItem);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Clear()
		{
			 Factory.ExecuteMethod(this, "Clear");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void DropDown()
		{
			 Factory.ExecuteMethod(this, "DropDown");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		/// <param name="pvargIndex">object pvargIndex</param>
		[SupportByVersion("MSForms", 2)]
		public void RemoveItem(object pvargIndex)
		{
			 Factory.ExecuteMethod(this, "RemoveItem", pvargIndex);
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Copy()
		{
			 Factory.ExecuteMethod(this, "Copy");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Cut()
		{
			 Factory.ExecuteMethod(this, "Cut");
		}

		/// <summary>
		/// SupportByVersion MSForms 2
		/// </summary>
		[SupportByVersion("MSForms", 2)]
		public void Paste()
		{
			 Factory.ExecuteMethod(this, "Paste");
		}

		#endregion

		#pragma warning restore
	}
}
