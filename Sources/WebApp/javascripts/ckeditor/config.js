﻿/*
Copyright (c) 2003-2010, CKSource - Frederico Knabben. All rights reserved.
For licensing, see LICENSE.html or http://ckeditor.com/license
*/

CKEDITOR.editorConfig = function( config )
{
	config.toolbar = 'Custom';
	
	config.toolbar_Custom =
    [
        ['Maximize', '-', 'Cut','Copy', 'PasteText','PasteFromWord','-', 'SpellChecker', 'Scayt'],
        ['Undo','Redo','-','Find','Replace','-','SelectAll','RemoveFormat'],
        '/',
        ['Bold','Italic','Underline'],
        ['NumberedList','BulletedList'],
        ['Image','Table','SpecialChar']
    ];
};