﻿// Copyright (c) Microsoft Corporation. All rights reserved.
// Licensed under the MIT license.

namespace teams_export.Models
{
    // Simple class to serialize user details
    public class CachedUser
    {
        public string DisplayName { get; set; }
        public string Email { get; set; }
        public string Avatar { get; set; }
    }
}