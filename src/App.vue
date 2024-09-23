<template>
  <div class="h-full p-2">
    <h1 class="text-2xl font-bold">Upload Data Penjualan</h1>
    <a
      href="#"
      class="inline-block mt-2 border-2 font-semibold border-[#1078CA] text-[#1078CA] p-1 rounded-md"
      >Unduh Template Data Penjualan</a
    >
    <p class="my-2">
      Pilih file Excel/CSV Anda di bawah ini untuk memuat data outlet
    </p>
    <div v-if="isUploading" class="text-center">
      <div class="mx-auto w-1/3">
        <div class="mt-3 mx-auto w-full bg-gray-200 rounded-full h-2.5">
          <div
            class="bg-[#3384F3] h-2.5 rounded-full"
            :style="{ width: uploadProgress + '%' }"
          ></div>
        </div>
        <div class="flex justify-end">
          <small class="text-sm text-slate-300">{{ uploadProgress }}%</small>
        </div>
      </div>
      <div class="mt-2">
        <h3 class="font-semibold">Proses Upload Data Sale</h3>
        <p class="mt-1 text-slate-300">
          Kami memerlukan waktu lebih untuk melakukan verifikasi data Anda.
        </p>
      </div>
    </div>
    <div
      v-if="file.length === 0 && !isUploading"
      class="flex flex-col gap-2 items-center border-2 border-slate-300 border-dashed py-3"
    >
      <p>Taruh file Excel/CSV di bawah ini</p>
      <small>Atau</small>
      <label
        for="file-sales"
        class="px-2 py-1 rounded-lg bg-[#1078CA] text-white"
        >Pilih File</label
      >
      <input
        class="hidden"
        id="file-sales"
        type="file"
        @change="onFileChange"
      />
    </div>
    <table
      v-if="file.length > 0 && !isUploading"
      class="mt-5 w-full text-sm text-left rtl:text-right text-gray-500 dark:text-gray-400"
    >
      <thead
        class="text-xs text-gray-700 uppercase bg-gray-50 dark:bg-gray-700 dark:text-gray-400"
      >
        <tr>
          <th scope="col" class="px-6 py-3">Order Date</th>
          <th scope="col" class="px-6 py-3">Region</th>
          <th scope="col" class="px-6 py-3">Manager</th>
          <th scope="col" class="px-6 py-3">Salesman</th>
          <th scope="col" class="px-6 py-3">Item</th>
          <th scope="col" class="px-6 py-3">Units</th>
          <th scope="col" class="px-6 py-3">Units Price</th>
          <th scope="col" class="px-6 py-3">Sale Amt</th>
        </tr>
      </thead>
      <tbody>
        <tr
          v-for="item in file"
          :key="item.id"
          class="bg-white border-b dark:bg-gray-800 dark:border-gray-700"
        >
          <td class="px-6 py-4">{{ item["OrderDate"] }}</td>
          <td class="px-6 py-4">{{ item["Region"] }}</td>
          <td class="px-6 py-4">{{ item["Manager"] }}</td>
          <td class="px-6 py-4">{{ item["SalesMan"] }}</td>
          <td class="px-6 py-4">{{ item["Item"] }}</td>
          <td class="px-6 py-4">{{ item["Units"] }}</td>
          <td class="px-6 py-4">{{ item["Unit_price"] }}</td>
          <td class="px-6 py-4">{{ item["Sale_amt"] }}</td>
        </tr>
      </tbody>
    </table>
  </div>
</template>

<script>
import * as XLSX from "xlsx/xlsx.mjs";

function readExcelData(file) {
  return new Promise((resolve, reject) => {
    const validMimeTypes = [
      "application/vnd.ms-excel",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ];

    const validExtensions = [".xls", ".xlsx"];

    const fileExtension = file.name
      .slice(file.name.lastIndexOf("."))
      .toLowerCase();

    if (
      !validMimeTypes.includes(file.type) ||
      !validExtensions.includes(fileExtension)
    ) {
      return reject(
        new Error(
          "Data yang Anda unggah tidak sesuai dengan template atau kurang lengkap."
        )
      );
    }

    const reader = new FileReader();

    reader.onload = (e) => {
      try {
        const loadStartTime = Date.now();
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const json = XLSX.utils.sheet_to_json(worksheet, {
          defval: null,
          header: 1,
        });

        let headers = json[0].map((header) => header.trim());

        let cleanData = json
          .slice(1)
          .map((row) => {
            let obj = {};
            row.forEach((value, index) => {
              if (value !== null && value !== undefined && value !== "") {
                obj[headers[index]] = value;
              }
            });
            return obj;
          })
          .filter((obj) => Object.keys(obj).length > 0);

        const loadEndTime = Date.now();
        const totalLoadTime = loadEndTime - loadStartTime;
        console.log(`Total loading time: ${totalLoadTime} milliseconds`);
        resolve({ cleanData, totalLoadTime });
      } catch (error) {
        reject(
          new Error(
            "Data yang Anda unggah tidak sesuai dengan template atau kurang lengkap."
          )
        );
      }
    };

    reader.onerror = (error) => {
      reject(error);
    };

    reader.readAsArrayBuffer(file);
  });
}

export default {
  data() {
    return {
      isUploading: false,
      file: [],
      uploadProgress: 0,
    };
  },

  methods: {
    onFileChange(e) {
      this.isUploading = true;
      const file = e.target.files[0];
      readExcelData(file)
        .then(({ cleanData, totalLoadTime }) => {
          for (let i = 0; i < cleanData.length; i++) {
            if (
              !cleanData[i].OrderDate ||
              !cleanData[i].Region ||
              !cleanData[i].Manager ||
              !cleanData[i].SalesMan ||
              !cleanData[i].Item ||
              !cleanData[i].Units ||
              !cleanData[i].Unit_price ||
              !cleanData[i].Sale_amt
            ) {
              // console.error("Data tidak sesuai dengan template");
              this.isUploading = false;
              this.$swal.fire({
                icon: "error",
                title: "Gagal Upload Ffile",
                text: "Data yang Anda unggah tidak sesuai dengan template atau kurang lengkap.",
              });
              return;
            }
          }

          this.file = cleanData;

          const intervalTime = 100; // Update setiap 100 ms
          const totalSteps = Math.ceil(totalLoadTime / intervalTime);
          const stepProgress = 100 / totalSteps;

          let currentStep = 0;
          // let interval = setInterval(() => {
          //   if (this.uploadProgress < 100) {
          //     this.uploadProgress += 10; // Menaikkan progress secara bertahap
          //   } else {
          //     clearInterval(interval);
          //     this.file = cleanData;
          //     this.isUploading = false; // Menghentikan upload saat progress mencapai 100%
          //   }
          // }, 100); // Simulasi progress setiap 100ms

          const interval = setInterval(() => {
            if (currentStep < totalSteps) {
              this.uploadProgress += stepProgress;
              currentStep++;
            } else {
              clearInterval(interval);
              // this.file = cleanData;
              this.isUploading = false; // Menghentikan upload saat progress mencapai 100%
            }
          }, intervalTime);

          // this.file = result;
          // this.isUploading = false;

          // console.log("Excel data:", result);
        })
        .catch((error) => {
          this.isUploading = false;
          this.$swal.fire({
            icon: "error",
            title: "Gagal Upload File",
            text: error.message,
          });
          // console.error("Error reading Excel data:", error);
        });
    },
  },
};
</script>
