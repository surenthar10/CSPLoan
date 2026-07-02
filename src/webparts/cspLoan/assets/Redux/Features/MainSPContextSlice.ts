import { createSlice } from "@reduxjs/toolkit";
import { IMainData } from "../../Config/interface";

const mainData: IMainData = {
  isAdmin: false,
  loanDetails: [],
};

const MainSPContext = createSlice({
  name: "MainSPContext",
  initialState: mainData,
  reducers: {
    setIsAdmin: (state, action) => {
      state.isAdmin = action.payload;
    },
    setLoanDetails: (state, action) => {
      state.loanDetails = action.payload;
    },
  },
});

export const { setIsAdmin, setLoanDetails } = MainSPContext.actions;

export default MainSPContext.reducer;
