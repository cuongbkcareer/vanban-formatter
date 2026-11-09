/// <reference types="office-js" />
import { webHost } from './web-host';
import { mountUI } from './ui';
import { wordHost } from './word-host';

/**
 * Điểm vào của khung bên cạnh trong Word.
 * Nếu trang được mở ngoài Word (ví dụ gõ thẳng địa chỉ vào trình duyệt) thì chạy như bản web độc lập.
 */
Office.onReady((info) => {
  const inWord = info.host === Office.HostType.Word;
  if (inWord && !Office.context.requirements.isSetSupported('WordApi', '1.3')) {
    document.body.dataset.host = 'unsupported';
    return;
  }
  mountUI(inWord ? wordHost : webHost);
}).catch(() => mountUI(webHost));
