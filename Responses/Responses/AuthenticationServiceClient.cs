﻿//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated by a tool.
//     Runtime Version:4.0.30319.1008
//
//     Changes to this file may cause incorrect behavior and will be lost if
//     the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

[System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
[System.ServiceModel.ServiceContractAttribute(ConfigurationName="IAuthenticationService")]
public interface IAuthenticationService
{
    
    [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IAuthenticationService/IsLoggedIn", ReplyAction="http://tempuri.org/IAuthenticationService/IsLoggedInResponse")]
    Checkbox.Wcf.Services.Proxies.ServiceOperationResultOfboolean IsLoggedIn();
    
    [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IAuthenticationService/Login", ReplyAction="http://tempuri.org/IAuthenticationService/LoginResponse")]
    Checkbox.Wcf.Services.Proxies.ServiceOperationResultOfstring Login(string userName, string password);
    
    [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IAuthenticationService/Logout", ReplyAction="http://tempuri.org/IAuthenticationService/LogoutResponse")]
    Checkbox.Wcf.Services.Proxies.ServiceOperationResultOfanyType Logout();
    
    [System.ServiceModel.OperationContractAttribute(Action="http://tempuri.org/IAuthenticationService/ValidateUser", ReplyAction="http://tempuri.org/IAuthenticationService/ValidateUserResponse")]
    Checkbox.Wcf.Services.Proxies.ServiceOperationResultOfboolean ValidateUser(string userName, string password);
}

[System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
public interface IAuthenticationServiceChannel : IAuthenticationService, System.ServiceModel.IClientChannel
{
}

[System.Diagnostics.DebuggerStepThroughAttribute()]
[System.CodeDom.Compiler.GeneratedCodeAttribute("System.ServiceModel", "4.0.0.0")]
public partial class AuthenticationServiceClient : System.ServiceModel.ClientBase<IAuthenticationService>, IAuthenticationService
{
    
    public AuthenticationServiceClient()
    {
    }
    
    public AuthenticationServiceClient(string endpointConfigurationName) : 
            base(endpointConfigurationName)
    {
    }
    
    public AuthenticationServiceClient(string endpointConfigurationName, string remoteAddress) : 
            base(endpointConfigurationName, remoteAddress)
    {
    }
    
    public AuthenticationServiceClient(string endpointConfigurationName, System.ServiceModel.EndpointAddress remoteAddress) : 
            base(endpointConfigurationName, remoteAddress)
    {
    }
    
    public AuthenticationServiceClient(System.ServiceModel.Channels.Binding binding, System.ServiceModel.EndpointAddress remoteAddress) : 
            base(binding, remoteAddress)
    {
    }
    
    public Checkbox.Wcf.Services.Proxies.ServiceOperationResultOfboolean IsLoggedIn()
    {
        return base.Channel.IsLoggedIn();
    }
    
    public Checkbox.Wcf.Services.Proxies.ServiceOperationResultOfstring Login(string userName, string password)
    {
        return base.Channel.Login(userName, password);
    }
    
    public Checkbox.Wcf.Services.Proxies.ServiceOperationResultOfanyType Logout()
    {
        return base.Channel.Logout();
    }
    
    public Checkbox.Wcf.Services.Proxies.ServiceOperationResultOfboolean ValidateUser(string userName, string password)
    {
        return base.Channel.ValidateUser(userName, password);
    }
}
