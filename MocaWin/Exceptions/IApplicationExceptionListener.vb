
Namespace Exceptions

    ''' <summary>
    ''' アプリケーションでキャッチしきれていない例外をキャッチした時に、
    ''' システム固有の処理を行うクラスを作成する為のインタフェース
    ''' </summary>
    ''' <remarks></remarks>
	Public Interface IApplicationExceptionListener

		''' <summary>
		''' アプリケーションでキャッチしきれていない例外が発生
		''' </summary>
		''' <param name="ex">対象の例外</param>
		''' <remarks></remarks>
		Sub ApplicationException(ByVal ex As Exception)

	End Interface

End Namespace
