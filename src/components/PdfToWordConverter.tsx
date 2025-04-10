import React, { useState } from 'react';
import * as docx from 'docx';
import { saveAs } from 'file-saver';
import * as pdfjsLib from 'pdfjs-dist';

// Khởi tạo PDF.js worker - cập nhật phiên bản để khớp với thư viện
pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js';

// Định nghĩa interface cho Document để TypeScript hiểu đúng kiểu dữ liệu
interface DocxDocumentInterface {
  sections: Array<{
    properties: any;
    children: any[];
  }>;
}

const PdfToWordConverter: React.FC = () => {
  const [pdfFile, setPdfFile] = useState<File | null>(null);
  const [isConverting, setIsConverting] = useState<boolean>(false);
  const [progress, setProgress] = useState<number>(0);
  const [error, setError] = useState<string>('');
  const [log, setLog] = useState<string[]>([]);

  const addLog = (message: string): void => {
    setLog(prevLog => [...prevLog, message]);
  };

  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>): void => {
    const file = e.target.files?.[0];
    if (file && file.type === 'application/pdf') {
      setPdfFile(file);
      setError('');
    } else {
      setPdfFile(null);
      setError('Vui lòng chọn tệp PDF hợp lệ.');
    }
  };

  const convertPdfToWord = async (): Promise<void> => {
    if (!pdfFile) {
      setError('Vui lòng chọn tệp PDF trước.');
      return;
    }

    try {
      setIsConverting(true);
      setProgress(0);
      setLog([]);
      addLog('Bắt đầu quá trình chuyển đổi...');

      // Đọc file PDF
      const arrayBuffer = await pdfFile.arrayBuffer();
      addLog('Đang tải tệp PDF...');
      
      // Tải PDF document
      const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
      const numPages = pdf.numPages;
      addLog(`PDF có ${numPages} trang. Đang xử lý...`);
      
      const paragraphs: docx.Paragraph[] = [];
      
      // Thêm tiêu đề tài liệu
      paragraphs.push(
        new docx.Paragraph({
          children: [
            new docx.TextRun({
              text: `Tài liệu chuyển đổi từ: ${pdfFile.name}`,
              bold: true,
              size: 28,
            }),
          ],
          alignment: docx.AlignmentType.CENTER,
        })
      );
      
      paragraphs.push(
        new docx.Paragraph({
          text: `Được tạo vào: ${new Date().toLocaleString()}`,
          alignment: docx.AlignmentType.CENTER,
        })
      );
      
      paragraphs.push(
        new docx.Paragraph({
          text: '',
          spacing: {
            after: 200,
          },
        })
      );
      
      // Xử lý từng trang của PDF
      for (let i = 1; i <= numPages; i++) {
        addLog(`Đang xử lý trang ${i}/${numPages}...`);
        setProgress(Math.floor((i / numPages) * 100));
        
        // Lấy trang
        const page = await pdf.getPage(i);
        
        // Cài đặt tỷ lệ để render
        const scale = 1.5;
        const viewport = page.getViewport({ scale });
        
        // Tạo canvas để render trang PDF
        const canvas = document.createElement('canvas');
        const context = canvas.getContext('2d');
        canvas.width = viewport.width;
        canvas.height = viewport.height;
        
        if (context) {
          // Render trang PDF vào canvas
          await page.render({
            canvasContext: context,
            viewport: viewport
          }).promise;
          
          // Chuyển đổi canvas thành dạng base64
          const imgData = canvas.toDataURL('image/jpeg', 0.95);
          
          // Thêm tiêu đề trang
          paragraphs.push(
            new docx.Paragraph({
              children: [
                new docx.TextRun({
                  text: `Trang ${i}`,
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: {
                before: 400,
                after: 200,
              },
              alignment: docx.AlignmentType.CENTER,
            })
          );
          
          // Tạo đối tượng hình ảnh từ base64
          const imageData = imgData.split(',')[1];
          
          if (!imageData) {
            addLog(`Cảnh báo: Không thể trích xuất dữ liệu hình ảnh cho trang ${i}`);
            continue;
          }
          
          // Thêm hình ảnh vào document
          try {
            const imagePararaph = new docx.Paragraph({
              children: [
                new docx.ImageRun({
                  data: Uint8Array.from(atob(imageData), c => c.charCodeAt(0)),
                  transformation: {
                    width: 600,
                    height: Math.floor(600 * (viewport.height / viewport.width)),
                  },
                }),
              ],
              alignment: docx.AlignmentType.CENTER,
            });
            
            paragraphs.push(imagePararaph);
          } catch (imageError) {
            addLog(`Cảnh báo: Không thể thêm hình ảnh cho trang ${i}: ${imageError instanceof Error ? imageError.message : 'Lỗi không xác định'}`);
          }
          
          // Thêm khoảng trống giữa các trang
          paragraphs.push(
            new docx.Paragraph({
              text: '',
              spacing: {
                after: 200,
              },
            })
          );
        }
      }
      
      // Tạo document Word mới với tất cả các đoạn văn bản đã tạo
      const doc = new docx.Document({
        sections: [
          {
            properties: {},
            children: paragraphs,
          },
        ],
      });
      
      addLog('Đang tạo tệp DOCX...');
      
      // Tạo blob để tải về
      const blob = await docx.Packer.toBlob(doc);
      
      // Tạo tên tệp Word từ tên tệp PDF
      const wordFileName = pdfFile.name.replace('.pdf', '') + '_converted.docx';
      
      // Tải tệp Word về
      saveAs(blob, wordFileName);
      
      addLog('Chuyển đổi hoàn tất! Tệp DOCX đã được tải về.');
      setIsConverting(false);
      setProgress(100);
      
    } catch (err) {
      console.error('Lỗi khi chuyển đổi:', err);
      setError(`Lỗi khi chuyển đổi: ${err instanceof Error ? err.message : 'Lỗi không xác định'}`);
      addLog(`Đã xảy ra lỗi: ${err instanceof Error ? err.message : 'Lỗi không xác định'}`);
      setIsConverting(false);
    }
  };

  return (
    <div className="flex flex-col items-center justify-center p-6 bg-gray-100 rounded-lg shadow-md">
      <h1 className="text-2xl font-bold mb-6 text-blue-700">Công cụ chuyển đổi PDF ảnh sang Word</h1>
      
      <div className="w-full max-w-md p-4 bg-white rounded-md shadow mb-6">
        <div className="mb-4">
          <label className="block text-gray-700 mb-2 font-medium">Chọn tệp PDF:</label>
          <input 
            type="file" 
            accept=".pdf" 
            onChange={handleFileChange} 
            className="w-full p-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500"
          />
        </div>
        
        {pdfFile && (
          <div className="mb-4">
            <p className="text-gray-700">
              <strong>Tệp đã chọn:</strong> {pdfFile.name} ({Math.round(pdfFile.size / 1024)} KB)
            </p>
          </div>
        )}
        
        {error && (
          <div className="mb-4 p-3 bg-red-100 text-red-700 rounded-md">
            {error}
          </div>
        )}
        
        <button 
          onClick={convertPdfToWord} 
          disabled={!pdfFile || isConverting} 
          className={`w-full py-2 px-4 rounded-md font-medium text-white ${!pdfFile || isConverting ? 'bg-gray-400 cursor-not-allowed' : 'bg-blue-600 hover:bg-blue-700'}`}
        >
          {isConverting ? 'Đang chuyển đổi...' : 'Chuyển đổi sang Word'}
        </button>
      </div>
      
      {isConverting && (
        <div className="w-full max-w-md mb-6">
          <div className="mb-2 flex justify-between">
            <span className="text-sm text-gray-600">Tiến trình:</span>
            <span className="text-sm font-medium">{progress}%</span>
          </div>
          <div className="w-full bg-gray-200 rounded-full h-2.5">
            <div className="bg-blue-600 h-2.5 rounded-full" style={{ width: `${progress}%` }}></div>
          </div>
        </div>
      )}
      
      {log.length > 0 && (
        <div className="w-full max-w-md p-4 bg-white rounded-md shadow h-48 overflow-auto">
          <h3 className="text-lg font-medium mb-2 text-gray-700">Nhật ký chuyển đổi:</h3>
          <div className="text-sm text-gray-600 font-mono">
            {log.map((entry, index) => (
              <div key={index} className="mb-1">
                {entry}
              </div>
            ))}
          </div>
        </div>
      )}
      
      <div className="mt-6 text-sm text-gray-600">
        <p className="mb-2"><strong>Lưu ý:</strong></p>
        <ul className="list-disc pl-5 space-y-1">
          <li>Công cụ này hoạt động tốt nhất với các tệp PDF có chứa hình ảnh quét.</li>
          <li>Tệp Word đầu ra sẽ chứa các hình ảnh được trích xuất từ PDF.</li>
          <li>Quá trình chuyển đổi diễn ra hoàn toàn trong trình duyệt, không có dữ liệu nào được gửi lên máy chủ.</li>
          <li>Đối với tệp PDF lớn, quá trình chuyển đổi có thể mất nhiều thời gian.</li>
        </ul>
      </div>
    </div>
  );
};

export default PdfToWordConverter; 