﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for SecuritySetting
/// </summary>
public class SecuritySetting
{
    /// <summary>
    /// Gets or sets a value indicating whether all pages will be forced to use SSL (no matter of a specified [NopHttpsRequirementAttribute] attribute)
    /// </summary>
    public bool ForceSslForAllPages { get; set; }

    /// <summary>
    /// Gets or sets an encryption key
    /// </summary>
    public string EncryptionKey { get; set; }

    /// <summary>
    /// Gets or sets a list of adminn area allowed IP addresses
    /// </summary>
    public List<string> AdminAreaAllowedIpAddresses { get; set; }

    /// <summary>
    /// Gets or sets a vaule indicating whether to hide admin menu items based on ACL
    /// </summary>
    public bool HideAdminMenuItemsBasedOnPermissions { get; set; }
}