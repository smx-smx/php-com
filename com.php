<?php
/**
 * Proof of Concept COM object written in PHP
 * Copyright(C) 2021 Stefano Moioli <smxdev4@gmail.com>
 */

define("S_OK", 0);
define("E_NOINTERFACE", 0x80004002);
define("E_NOTIMPL", 0x80004001);
define("IID_IUnknown"     , "00000000-0000-0000-c000-000000000046");
define("IID_IClassFactory", "00000001-0000-0000-c000-000000000046");

use FFI\CData;

function makeGuid(FFI $ffi, string $guidStr) : CData {
	$str = str_replace('-', '', $guidStr);
	$p = 0;

	$g = $ffi->new('GUID');
	$g->Data1 = hexdec(substr($str, $p, 8)); 	$p += 8;
	$g->Data2 = hexdec(substr($str, $p, 4));	$p + 4;
	$g->Data3 = hexdec(substr($str, $p, 4));	$p + 4;
	for($i=0; $i<8; $i++){
		$g->Data4[$i] = hexdec(substr($str, $p, 2)); $p += 2;
	}
	return $g;
}

function makeVtb(FFI $ffi, array $functions) : CData {
	$fieldMap = array();

	$cdef = "struct {";
	foreach($functions as $fnEntry){
		list($name, $argTypes, $callable) = $fnEntry;
		$cdef.= "long (__stdcall *{$name})({$argTypes});";
		$fieldMap[$name] = $callable;
	}
	$cdef .= "}";

	$structT = $ffi->type($cdef);
	
	//$vtb = $ffi->new($structT, false, true);
	$vtb = $ffi->new($structT);
	foreach($fieldMap as $name => $fn){
		$vtb->{$name} = $fn;
	}

	return $vtb;
}

function makeCComPtr(CData $vtb){
	$vtbPtr = FFI::addr($vtb);
	$vtbPtrT = FFI::typeof($vtbPtr);

	$ccomPtrT = FFI::arrayType($vtbPtrT, [1]);
	$ccomPtr = FFI::new($ccomPtrT);
	$ccomPtr[0] = $vtbPtr;

	return $ccomPtr;
}

function registerClass(FFI $ffi, string $guidStr, array $functions){
	$guid = makeGuid($ffi, $guidStr);
	$guidPtr = FFI::addr($guid);

	$cookie = $ffi->new("uint32_t");
	$cookiePtr = FFI::addr($cookie);

	$vtb = makeVtb($ffi, $functions);
	$ccomPtr = makeCComPtr($vtb);

	$ret = $ffi->CoRegisterClassObject(
		$guidPtr,
		$ccomPtr,
		0x1, 0x1,
		$cookiePtr
	);
	assert($ret === S_OK);
}

function createInstance(FFI $ffi, string $classGuidStr, string $ifaceGuidStr, array $functions) : CData {
	$clsid = makeGuid($ffi, $classGuidStr);
	$iid = makeGuid($ffi, $ifaceGuidStr);

	$refClsid = FFI::addr($clsid);
	$refIID = FFI::addr($iid);

	$vtb = makeVtb($ffi, $functions);
	$ccomPtr = makeCComPtr($vtb);
	$ppv = FFI::addr($ccomPtr);

	$ret = $ffi->CoCreateInstance(
		$refClsid,
		null,
		0x1,
		$refIID,
		$ppv
	);
	
	assert($ret === S_OK);
	return $ppv;
}

function trace(){
	$bt = debug_backtrace(DEBUG_BACKTRACE_PROVIDE_OBJECT, 2);
	$caller = array_pop($bt);
	print("[{$caller['class']}:{$caller['function']}:{$caller['line']}]\n");
}

function nativeGuidStr(CData $native, bool $isPtr = true){
	$g = ($isPtr) ? $native[0] : $native;
	return sprintf("%08x-%04x-%04x-%02x%02x-%02x%02x%02x%02x%02x%02x",
		$g->Data1, $g->Data2, $g->Data3,
		$g->Data4[0], $g->Data4[1], $g->Data4[2], $g->Data4[3],
		$g->Data4[4], $g->Data4[5], $g->Data4[6], $g->Data4[7]
	);
}

$ole32 = FFI::load("ole32.h");
$ole32->CoInitializeEx(null, 0x2);

abstract class IUnknown {
	private $refCount = 0;

	public function AddRef($pThis){ trace(); return ++$this->refCount; }
	public function Release($pThis){ trace(); return --$this->refCount; }
}

class MyFactory extends IUnknown {
	const GUID = "d6d16ab8-f65f-4152-8cf6-4f5e00e8aaa7";

	public function QueryInterface($pThis, $riid, $ppvObject){
		trace();
		$guid = nativeGuidStr($riid);
		switch($guid){
			case IID_IUnknown:
			case IID_IClassFactory:
				$ppvObject[0] = $pThis;
				$this->AddRef($pThis);
				return S_OK;
		}

		return E_NOINTERFACE;
	}

	public function CreateInstance($pThis, $pUnkOuter, $iid, $ppv){
		trace();
		assert($pUnkOuter === null);

		$obj = new HelloObj();
		$hr = $obj->QueryInterface($pThis, $iid, $ppv);
		$obj->Release(null);

		return $hr;
	}

	public function LockServer($pThis, $bLock){
		return E_NOTIMPL;
	}
}

class HelloObj extends IUnknown {
	const CLSID_HELLO = 'd6d16ab8-f65f-4152-8cf6-4f5e00e8aaa7';
	const IID_HELLO   = '6531d857-c22f-4add-b2d5-e9785e39fc46';

	public function __construct(){}

	public function QueryInterface($pThis, $riid, $ppvObject){
		trace();
		$guid = nativeGuidStr($riid);
		switch($guid){
			case IID_IUnknown:
			case HelloObj::IID_HELLO:
				$ppvObject[0] = $pThis;
				$this->AddRef($pThis);
				return S_OK;
		}

		return E_NOINTERFACE;
	}

	public function Hello($pThis){
		trace();
		print("HELLO, WORLD!!\n");
		return S_OK;
	}
}

function makeIUnknown($obj){
	return array(
		['QueryInterface', 'void *This, REFIID riid, void **ppvObject', array($obj, 'QueryInterface')],
		['AddRef', 'void *This', array($obj, 'AddRef')],
		['Release', 'void *This', array($obj, 'Release')]
	);
}

$factory = new MyFactory();


registerClass($ole32, HelloObj::CLSID_HELLO, array_merge(
	makeIUnknown($factory),
	array(
		['CreateInstance', 'void *This, void *pUnkOuter, REFIID riid, void **ppvObject', array($factory, 'CreateInstance')],
		['LockServer', 'void *This, uint32_t bLock', array($factory, 'LockServer')]
	)
));

$obj = new HelloObj();
$pHello = createInstance(
	$ole32,
	HelloObj::CLSID_HELLO,
	HelloObj::IID_HELLO,
	array_merge(
		makeIUnknown($factory),
		array(
			['Hello', 'void *This', array($obj, 'Hello')]
		)
	)
);

$pvtbl = $pHello[0];
$vtbl = $pvtbl[0];
$pfnHello = $vtbl->Hello;
$pfnHello($pHello);

//$ole32->CoUninitialize();