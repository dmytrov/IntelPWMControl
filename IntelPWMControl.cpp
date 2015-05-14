// this file is in the public domain

#include <windows.h>
#include <objbase.h>
#include <stdio.h>

#undef NDEBUG
#include <assert.h>

struct ICUIPower : public IUnknown
{
  virtual HRESULT STDMETHODCALLTYPE IsSupported(LPBOOL pSupported);
  virtual HRESULT STDMETHODCALLTYPE dummy1();
  virtual HRESULT STDMETHODCALLTYPE dummy2();
  virtual HRESULT STDMETHODCALLTYPE dummy3();
  virtual HRESULT STDMETHODCALLTYPE dummy4();
  virtual HRESULT STDMETHODCALLTYPE GetPWMFrequency(LPUINT puiInverterType, LPDWORD pdwPWMFreq, LPDWORD pdwErrorCodes);
  virtual HRESULT STDMETHODCALLTYPE SetPWMFrequency(UINT uiInverterType, DWORD dwPWMFreq, LPDWORD pdwErrorCodes);
};

int CALLBACK WinMain(HINSTANCE hInstance, HINSTANCE hPrevInstance, LPSTR lpCmdLine, int nCmdShow)
{
  int targetPWMFreq = 500;
  sscanf_s(lpCmdLine, "%d", &targetPWMFreq);
  targetPWMFreq = max(targetPWMFreq, 50);
  targetPWMFreq = min(targetPWMFreq, 10000);
  assert(SUCCEEDED(CoInitialize(0)));
  CLSID clsid = { 0xC332C124, 0x340D, 0x4430, { 0xAA, 0x0D, 0xC7, 0x56, 0x02, 0x87, 0x6F, 0xCC } };
  IID IID_ICUIPower = { 0x299D88F9, 0x2CBD, 0x4225, { 0xBF, 0x19, 0xFC, 0xD1, 0x64, 0xC5, 0x4C, 0x3F } };
  IUnknown *pUnknown;
  assert(SUCCEEDED(CoCreateInstance(clsid, 0, CLSCTX_LOCAL_SERVER, IID_IUnknown, (void**)&pUnknown)));
  ICUIPower *pCUIPower;
  assert(SUCCEEDED(pUnknown->QueryInterface(IID_ICUIPower, (void**)&pCUIPower)));
  while (1)
  {
    UINT uiInverterType = 0;
    DWORD dwPWMFreq = 0,
          dwErrorCodes = 0;
    assert(SUCCEEDED(pCUIPower->GetPWMFrequency(&uiInverterType, &dwPWMFreq, &dwErrorCodes)));
    if (dwPWMFreq != targetPWMFreq)
    {
      dwPWMFreq = targetPWMFreq;
      assert(SUCCEEDED(pCUIPower->SetPWMFrequency(uiInverterType, dwPWMFreq, &dwErrorCodes)));
    }
    Sleep(1000);
  }
  pCUIPower->Release();
  pUnknown->Release();
  CoUninitialize();
  return 0;
}

