import * as React from "react";
import { useEffect, useState } from "react";
import styles from "../BulkUpload/BulkUpload.module.scss";

const BulkUpload = () => {
  const [activeStep, setActiveStep] = useState<number>(1);

  useEffect(() => {
    setActiveStep(1);
  }, []);
  return (
    <div className={styles.bulkUploadWrapper}>
      <div className={styles.stepper}>
        <div
          className={`${styles.step} ${activeStep === 1 ? styles.active : ""}`}
        >
          <div className={styles.circle}>1</div>
          <span>Select & Map Files</span>
        </div>

        <div className={styles.connector}></div>

        <div
          className={`${styles.step} ${activeStep === 2 ? styles.active : ""}`}
        >
          <div className={styles.circle}>2</div>
          <span>Assign Tags</span>
        </div>

        <div className={styles.connector}></div>

        <div
          className={`${styles.step} ${activeStep === 3 ? styles.active : ""}`}
        >
          <div className={styles.circle}>3</div>
          <span>Review & Submit</span>
        </div>
      </div>

      {activeStep === 1 && (
        <div className={styles.uploadArea}>
          <i className={`pi pi-upload ${styles.uploadIcon}`}></i>

          <div className={styles.uploadTitle}>
            Click or drag files here to upload
          </div>

          <div className={styles.uploadText}>
            Supports up to 70+ files — path auto-suggested from filename prefix
          </div>
        </div>
      )}
    </div>
  );
};

export default BulkUpload;
