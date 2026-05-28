import sys
import os
import ctypes
import win32com.client as win32

def show_message(title, text, style=0):
    # style: 0 = OK, 16 = Critical, 48 = Warning, 64 = Info
    return ctypes.windll.user32.MessageBoxW(0, text, title, style)

def convert_to_pdf(hwp_path):
    if not os.path.exists(hwp_path):
        show_message("오류", f"파일을 찾을 수 없습니다:\n{hwp_path}", 16)
        return
    
    # Generate pdf file path
    dir_name = os.path.dirname(hwp_path)
    base_name = os.path.splitext(os.path.basename(hwp_path))[0]
    pdf_path = os.path.join(dir_name, base_name + ".pdf")
    
    hwp = None
    try:
        # Launch Hancom Office (Try EnsureDispatch first, fallback to Dispatch)
        try:
            hwp = win32.gencache.EnsureDispatch("HWPFrame.HwpObject")
        except Exception:
            hwp = win32.Dispatch("HWPFrame.HwpObject")
        
        # Hide the window immediately
        hwp.XHwpWindows.Item(0).Visible = False
        
        # Register the security module (it will bypass the warning dialog)
        hwp.RegisterModule("FilePathCheckDLL", "FilePathCheckerModuleExample")
        
        # Open file
        # We pass format as empty string so it auto-detects HWP/HWPX
        # We pass forceopen:true to force open
        opened = hwp.Open(hwp_path, "", "forceopen:true")
        if not opened:
            show_message("오류", f"한글 파일을 열 수 없습니다.\n파일 경로: {hwp_path}", 16)
            return
            
        # Save as PDF
        # Format "PDF" converts the document to PDF
        saved = hwp.SaveAs(pdf_path, "PDF")
        if not saved:
            show_message("오류", "PDF 저장에 실패했습니다.", 16)
            return
            
    except Exception as e:
        show_message("오류", f"변환 중 오류가 발생했습니다:\n{str(e)}", 16)
    finally:
        if hwp is not None:
            try:
                hwp.Quit()
            except Exception:
                pass

if __name__ == "__main__":
    if len(sys.argv) < 2:
        show_message("정보", "이 프로그램은 한글 파일을 PDF로 변환하는 백그라운드 서비스입니다.\n\n사용법: converter.py [한글 파일 경로]", 64)
        sys.exit(0)
        
    hwp_path = os.path.abspath(sys.argv[1])
    convert_to_pdf(hwp_path)
