using System;
using System.Collections.Generic;

namespace ExcelHandler.Models
{
    public class Error
    {
        public string Message { get; set; }
        public int Row { get; set; }
    }

    public class ExcelResponse<T>
    {
        public int Code { get; set; }

        public List<Error> Errors { get; set; }

        public T Data { get; set; }

        public static ExcelResponse<T> GetResult(int code, List<Error> errors, T data = default(T))
        {
            return new ExcelResponse<T>
            {
                Code = code,
                Errors = errors,
                Data = data
            };
        }

        public static ExcelResponse<T> GetError(string errorMessage)
        {
            return new ExcelResponse<T> { Code = -1, Errors = new List<Error> { new Error { Row = 0, Message = errorMessage } } };

        }
    }
}