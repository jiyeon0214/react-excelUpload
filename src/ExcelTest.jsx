import { useState } from 'react';
import * as XLSX from 'xlsx';

const ExcelTest = () => {
    const [data, setData] = useState([]);
    const headers = data.length > 0 ? Object.keys(data[0]) : [];

    // 병합 셀에 값 채워넣기
    function fillMergedCells(worksheet) {
        const merges = worksheet['!merges'] || []; // worksheet['!merges'] : 병합된 셀 범위 목록을 담고 있는 배열(start ~ end)

        merges.forEach(({ s, e }) => {
            const startCellAddr = XLSX.utils.encode_cell(s);
            const originalCell = worksheet[startCellAddr];

            if (!originalCell) return;

            for (let r = s.r; r <= e.r; r++) {
                for (let c = s.c; c <= e.c; c++) {
                    const cellAddr = XLSX.utils.encode_cell({ c, r });

                    // 셀이 비어 있을 경우에만 복사
                    if (!worksheet[cellAddr]) worksheet[cellAddr] = {};

                    // 값 복사
                    worksheet[cellAddr].v = originalCell.v;

                    // 타입 복사 (문자열, 숫자, 날짜 등)
                    worksheet[cellAddr].t = originalCell.t;

                    // 날짜 등의 포맷 정보가 있으면 복사
                    if (originalCell.z) worksheet[cellAddr].z = originalCell.z;

                    // 수식이 있으면 복사
                    if (originalCell.f) worksheet[cellAddr].f = originalCell.f;

                    // 포맷된 텍스트(w)도 있으면 복사 (sheet_to_html 같은 경우에 사용)
                    if (originalCell.w) worksheet[cellAddr].w = originalCell.w;
                }
            }
        });
    }

    const uploadFile = (e) => {
        const file = e.target.files[0]; // FileList(유사 배열 객체) 반환됨
        const reader = new FileReader(); // 브라우저 내장 객체. input type="file"으로 업로드한 파일을 읽을 수 있게 해 줌

        /* 
        파일 읽기 성공 시 onload, 실패 시 onerror 이벤트 발생
        reader.onload 핸들러는 반드시 readAsArrayBuffer 호출 전에 등록해야 함.
        (파일 읽기가 매우 빠르게 완료될 수 있어서, 핸들러가 등록되기 전에 이벤트가 발생하면 놓칠 수 있기 때문)
        */
        reader.onload = (ev) => {
            const fileArray = ev.target.result; // reader.readAsArrayBuffer(file) -> ArrayBuffer(이진 데이터 객체)
            const workbook = XLSX.read(fileArray, {
                type: 'array',
                cellDates: true, // 날짜 셀 Date 객체로 변환
                dateNF: 'yyyy-mm-dd', // 날짜 포맷 서식
            }); // ArrayBuffer를 엑셀 문서 객체(workbook)로 변환

            const worksheet = workbook.Sheets[workbook.SheetNames[0]]; // 시트 인덱스 선택
            fillMergedCells(worksheet); // 병합된 셀에 데이터 복사

            const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '', raw: false }); // json 배열로 변환. { defval: '' } : 비어 있는 셀을 빈 문자열로 처리(기본 설정은 undefined)
            // console.log(jsonData);
            setData(jsonData);
        };

        reader.readAsArrayBuffer(file); // 파일 읽기가 시작됨 (비동기)
    };

    const onCellChange = (row, col, value) => {
        const updated = [...data];
        updated[row][col] = value;
        setData(updated);
    };

    const save = () => {
        const newData = JSON.stringify(data, null, 2);
        console.log(newData);
    };

    return (
        <div>
            <input type="file" id="file" name="file" accept=".xlsx, .xls" onChange={uploadFile} />
            <table>
                <thead>
                    <tr>
                        {headers.map((header) => (
                            <th key={header}>{header}</th>
                        ))}
                    </tr>
                </thead>
                <tbody>
                    {data.map((row, rowIdx) => (
                        <tr key={rowIdx}>
                            {headers.map((key) => (
                                <td key={key}>
                                    <input
                                        type="text"
                                        value={row[key]}
                                        onChange={(e) => onCellChange(rowIdx, key, e.target.value)}
                                    />
                                </td>
                            ))}
                        </tr>
                    ))}
                </tbody>
            </table>
            <button onClick={save}>저장</button>
        </div>
    );
};

export default ExcelTest;
