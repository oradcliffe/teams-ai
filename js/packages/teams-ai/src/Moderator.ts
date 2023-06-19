/**
 * @module teams-ai
 */
/**
 * Copyright (c) Microsoft Corporation. All rights reserved.
 * Licensed under the MIT License.
 */

import { TurnContext } from 'botbuilder-core';
import { ConfiguredAIOptions } from './AI';
import { Plan } from './Planner';
import { PromptTemplate } from './Prompts';
import { TurnState } from './TurnState';

/**
 * A moderator is responsible for reviewing and approving AI prompts and plans.
 *
 * @template TState Type of the applications turn state.
 */
export interface Moderator<TState extends TurnState> {
    /**
     * Reviews an incoming utterance and generated prompt before it's sent to the planner.
     *
     * @remarks
     * The moderator can review the incoming utterance for things like prompt injection attacks
     * or the leakage of sensitive information. The moderator can also review the generated prompt
     * to ensure it's appropriate for the current conversation.
     *
     * To approve a prompt, simply return undefined. Returning a new plan bypasses the planner and
     * redirects to a new set of actions. Typically the moderator will return a new plan with a
     * single DO command that calls `AI.FlaggedInputActionName` to flag the input for review.
     *
     * The moderator can pass any entities that make sense to the redirected action.
     * @param context Context for the current turn of conversation.
     * @param state Application state for the current turn of conversation.
     * @param prompt Generated prompt to be reviewed.
     * @param options Current options for the AI system.
     * @returns An undefined value to approve the prompt or a new plan to redirect to if not approved.
     */
    reviewPrompt(
        context: TurnContext,
        state: TState,
        prompt: PromptTemplate,
        options: ConfiguredAIOptions<TState>
    ): Promise<Plan | undefined>;

    /**
     * Reviews a plan generated by the planner before its executed.
     *
     * @remarks
     * The moderator can review the plan to ensure it's appropriate for the current conversation.
     *
     * To approve a plan simply return the plan that was passed in. A new plan can be returned to
     * redirect to a new set of actions. Typically the moderator will return a new plan with a
     * single DO command that calls `AI.FlaggedOutputActionName` to flag the output for review.
     *
     * The moderator can pass any entities that make sense to the redirected action.
     * @param context Context for the current turn of conversation.
     * @param state Application state for the current turn of conversation.
     * @param plan Plan generated by the planner.
     * @returns The plan to execute. Either the current plan passed in for review or a new plan.
     */
    reviewPlan(context: TurnContext, state: TState, plan: Plan): Promise<Plan>;
}
