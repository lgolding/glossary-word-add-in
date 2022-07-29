// Avoid the error "ReferenceError: regeneratorRuntime is not defined".
// https://stackoverflow.com/questions/42535270/regeneratorruntime-is-not-defined-when-running-jest-test/57439821#57439821
import "regenerator-runtime/runtime";
