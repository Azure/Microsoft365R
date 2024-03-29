% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/ms_team_member.R
\docType{class}
\name{ms_team_member}
\alias{ms_team_member}
\title{Teams/channel member}
\format{
An R6 object of class \code{ms_team_member}, inheriting from \code{ms_object}.
}
\description{
Class representing a member of a team or channel (which will normally be a user in Azure Active Directory).
}
\section{Fields}{

\itemize{
\item \code{token}: The token used to authenticate with the Graph host.
\item \code{tenant}: The Azure Active Directory tenant for the parent object.
\item \code{type}: One of "team member", "channel member" or "chat member" depending on the parent object.
\item \code{properties}: The item properties (metadata).
}
}

\section{Methods}{

\itemize{
\item \code{new(...)}: Initialize a new object. Do not call this directly; see 'Initialization' below.
\item \code{delete(confirm=TRUE)}: Delete this member.
\item \code{update(...)}: Update the member's properties (metadata) in Microsoft Graph.
\item \code{do_operation(...)}: Carry out an arbitrary operation on the member.
\item \code{sync_fields()}: Synchronise the R object with the member metadata in Microsoft Graph.
\item \code{get_aaduser()}: Get the AAD information for the member; returns an object of class \code{\link[AzureGraph:az_user]{AzureGraph::az_user}}.
}
}

\section{Initialization}{

Creating new objects of this class should be done via the \code{get_member} and \code{list_members} methods of the \code{\link{ms_team}}and \code{\link{ms_channel}} classes. Calling the \code{new()} method for this class only constructs the R object; it does not call the Microsoft Graph API to retrieve or create the actual member.
}

\seealso{
\code{\link{ms_team}}, \code{\link{ms_channel}}

\href{https://learn.microsoft.com/en-us/graph/overview}{Microsoft Graph overview},
\href{https://learn.microsoft.com/en-us/graph/api/resources/teams-api-overview?view=graph-rest-1.0}{Microsoft Teams API reference}
}
