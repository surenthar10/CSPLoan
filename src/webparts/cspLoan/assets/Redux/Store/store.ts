import { configureStore } from "@reduxjs/toolkit";
import MainSPContextReducer from "../Features/MainSPContextSlice";

const store = configureStore({
  reducer: {
    MainSPContext: MainSPContextReducer,
  },
  devTools: false,
});

export default store;
